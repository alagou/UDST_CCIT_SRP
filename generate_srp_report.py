"""Build the CCIT SRP project-status report.

Run from the project folder:

    python generate_srp_report.py

Forms with a share_url are downloaded from SharePoint first.
No arguments are required for a normal daily run.
"""

from __future__ import annotations

import logging
import sys
import urllib.request
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd

from config import CONFIG, ROOT, iter_form_configs, project_path
from html_dashboard import build_html
from id_registry import ProjectRegistry
from merge import build_dataset, build_view, write_master
from validation import add_cross_checks, add_issue, count_by_severity, write_validation_report

KNOWN_ROLES = {"projects", "memberships", "registrations", "status_overrides"}


def fetch_latest_forms(config: dict) -> None:
    """Download each form that has a SharePoint ``share_url`` into ``input_folder``.

    Sharing links of the form ``:x:/g/personal/.../<id>`` are converted to
    the OneDrive download address. A failed download leaves the previous
    local workbook in place so the report can still run from that copy.
    """
    folder = project_path(config["input_folder"])
    folder.mkdir(parents=True, exist_ok=True)
    for name, form_cfg in iter_form_configs(config):
        share_url = str(form_cfg.get("share_url") or "").strip()
        if not share_url:
            continue
        try:
            saved = _download_share(share_url, folder)
        except Exception as exc:
            logging.error("Could not download %s from SharePoint: %s", name, exc)
            continue
        logging.info("Downloaded %s to %s", name, saved.name)


def main() -> int:
    _configure_stdio()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    config = CONFIG
    fetch_latest_forms(config)
    issues: list = []
    frames = {role: [] for role in KNOWN_ROLES}
    input_folder = project_path(config["input_folder"])
    input_folder.mkdir(parents=True, exist_ok=True)

    for name, form_cfg in iter_form_configs(config):
        role = form_cfg.get("role")
        if role not in KNOWN_ROLES:
            add_issue(issues, "ERROR", "", name, f"Unknown form role '{role}'", "")
            continue
        frame = _load_form(name, form_cfg, input_folder, issues)
        if frame is not None:
            frames[role].append(frame)

    faculty = _concat(frames["projects"])
    student = _concat(frames["memberships"])
    registration = _concat(frames["registrations"])
    admin = _concat(frames["status_overrides"])
    if student is None and registration is None:
        logging.info("No student registration export yet. Membership stays on the faculty name/ID list.")
    if admin is None:
        logging.info("Admin Form not found. Status will be computed from capacity and membership.")

    registry = ProjectRegistry(project_path(config["registry_path"]), config["academic_year"])
    projects, memberships, merge_issues, created, reused = build_dataset(
        faculty, student, admin, registry, config, registration
    )
    issues.extend(merge_issues)
    add_cross_checks(projects, memberships, registry.duplicate_project_ids(), issues, config)
    _write_file(registry.path, registry.save)
    logging.info("Project IDs: %s new, %s reused.", created, reused)

    cards, summary, programs = build_view(projects, issues)
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    output_dir = project_path(config["output_folder"])
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "html": output_dir / config["outputs"]["html"],
        "master": output_dir / config["outputs"]["master"],
        "validation": output_dir / config["outputs"]["validation"],
    }
    lookup_path = project_path(config["lookup_path"])
    _write_file(paths["validation"], lambda: write_validation_report(paths["validation"], issues))
    _write_file(paths["master"], lambda: write_master(paths["master"], projects, memberships))
    _write_file(paths["html"], lambda: build_html(paths["html"], cards, summary, programs, generated_at))
    _write_file(
        lookup_path,
        lambda: registry.export_lookup(
            lookup_path,
            {project["project_id"]: project["supervisor_name"] for project in projects},
        ),
    )
    _print_summary(summary, issues, paths, registry)
    return 0


def _write_file(path, writer) -> None:
    """Write one output. A file left open in Excel is reported and skipped."""
    try:
        writer()
    except PermissionError:
        logging.error("Close %s in the other program, then run this again.", path)


def _download_share(share_url: str, dest_dir: Path) -> Path:
    request = urllib.request.Request(
        _sharing_download_url(share_url),
        headers={"User-Agent": "Mozilla/5.0"},
    )
    with urllib.request.urlopen(request, timeout=60) as response:
        payload = response.read()
        disposition = response.headers.get("Content-Disposition", "")
        content_type = response.headers.get("Content-Type", "")
    if not payload.startswith(b"PK"):
        raise ValueError(f"SharePoint did not return an Excel file ({content_type or 'unknown type'})")
    filename = _filename_from_disposition(disposition)
    dest = dest_dir / filename
    dest.write_bytes(payload)
    return dest


def _sharing_download_url(share_url: str) -> str:
    parsed = urlparse(share_url.strip())
    if "download.aspx" in parsed.path:
        return share_url.strip()
    parts = [part for part in parsed.path.split("/") if part]
    try:
        marker = parts.index(":x:")
    except ValueError as exc:
        raise ValueError(f"Unrecognized SharePoint sharing link: {share_url}") from exc
    rest = parts[marker + 1 :]
    if rest and rest[0] == "g":
        rest = rest[1:]
    if len(rest) < 2:
        raise ValueError(f"Unrecognized SharePoint sharing link: {share_url}")
    token = rest[-1].split("?")[0]
    site = "/".join(rest[:-1])
    return f"{parsed.scheme}://{parsed.netloc}/{site}/_layouts/15/download.aspx?share={token}"


def _filename_from_disposition(disposition: str) -> str:
    message = EmailMessage()
    message["Content-Disposition"] = disposition or "attachment"
    filename = Path(message.get_filename() or "download.xlsx").name
    filename = filename.replace("\x00", "").strip() or "download.xlsx"
    if not filename.lower().endswith(".xlsx"):
        filename = f"{filename}.xlsx"
    return filename


def _load_form(name: str, form_cfg: dict, input_folder, issues: list):
    path = _discover(name, form_cfg, input_folder, issues)
    if path is None:
        return None
    sheet = form_cfg.get("sheet") or "Sheet1"
    try:
        frame = pd.read_excel(path, sheet_name=sheet)
    except Exception as exc:
        add_issue(issues, "ERROR", "", name, "Could not read form workbook", f"{path.name}: {exc}")
        logging.error("Could not read %s (%s): %s", name, path, exc)
        return None
    frame = _stringify(frame)
    missing = [column for column in form_cfg.get("required_columns", []) if column not in frame.columns]
    if missing:
        found = " | ".join(str(column) for column in frame.columns)
        for column in missing:
            add_issue(
                issues,
                "ERROR",
                "",
                column,
                f"Required column missing from {path.name}",
                found,
            )
        logging.error("%s is missing required columns: %s", path.name, ", ".join(missing))
        return None
    mapped = {}
    for source, dest in form_cfg.get("column_map", {}).items():
        if source in frame.columns:
            mapped[dest] = frame[source]
    logging.info("Loaded %s row(s) from %s", len(frame), path)
    if not mapped:
        return pd.DataFrame()
    return pd.DataFrame(mapped)


def _discover(name: str, form_cfg: dict, input_folder, issues: list):
    pattern = form_cfg.get("glob") or "*.xlsx"
    matches = [
        candidate
        for candidate in input_folder.glob(pattern)
        if candidate.is_file() and not candidate.name.startswith("~$")
    ]
    matches.sort(key=lambda candidate: candidate.stat().st_mtime, reverse=True)
    if len(matches) > 1:
        logging.warning(
            "%s matched %s files in %s; using the newest (%s).",
            name,
            len(matches),
            input_folder,
            matches[0].name,
        )
    if matches:
        return matches[0]
    fallback = project_path(form_cfg.get("path", ""))
    if fallback.exists():
        return fallback
    message = f"Form file not found in {input_folder} (glob {pattern})"
    if form_cfg.get("optional", False):
        logging.info("%s: %s", name, message)
        add_issue(issues, "INFO", "", name, "Optional form file not found", message)
    else:
        logging.error("%s: %s", name, message)
        add_issue(issues, "ERROR", "", name, "Required form file not found", message)
    return None


def _stringify(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.copy()
    frame.columns = [str(column).strip() for column in frame.columns]
    for column in frame.columns:
        frame[column] = frame[column].map(_stringify_cell)
    return frame


def _stringify_cell(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return str(value).strip()


def _concat(frames: list):
    if not frames:
        return None
    return pd.concat(frames, ignore_index=True)


def _print_summary(summary: dict, issues: list, paths: dict, registry: ProjectRegistry) -> None:
    counts = count_by_severity(issues)
    print("")
    print("CCIT SRP report complete")
    print(f"Projects: {summary['total_projects']}")
    print(f"Active projects: {summary['projects_active']}")
    print(f"Students assigned: {summary['total_students_assigned']}")
    print(f"Recruiting: {summary['projects_recruiting']}")
    print(f"Team full: {summary['projects_full']}")
    print(f"No students: {summary['projects_with_no_students']}")
    print(f"Available positions: {summary['total_available_positions']}")
    print(f"Requiring attention: {summary['projects_requiring_attention']}")
    print(
        f"Validation: {counts['ERROR']} ERROR, {counts['WARNING']} WARNING, {counts['INFO']} INFO"
    )
    print("Outputs:")
    for label, path in paths.items():
        print(f"  {label}: {path.resolve()}")
    print(f"  registry: {registry.path.resolve()}")
    print(f"  lookup: {project_path(CONFIG['lookup_path']).resolve()}")


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        logging.exception("Report generation failed: %s", exc)
        raise SystemExit(1)
