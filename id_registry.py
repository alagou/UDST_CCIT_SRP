"""Persistent Project_ID registry.

IDs look like ``SRP26-001``. A project keeps its ID across runs as long as
the normalized supervisor email and project title stay the same, even when
the Forms export is reordered.
"""

from __future__ import annotations

import re
from collections import Counter
from datetime import date
from pathlib import Path

import pandas as pd

REGISTRY_COLUMNS = [
    "Project_ID",
    "Supervisor_Email",
    "Project_Title",
    "First_Seen_Response_Id",
    "Date_Assigned",
]


def match_key(email, title) -> str:
    """Build the stable key: normalized email + normalized title."""

    def norm(value) -> str:
        text = "" if value is None else str(value)
        return re.sub(r"\s+", " ", text).strip().lower()

    return f"{norm(email)}|{norm(title)}"


class ProjectRegistry:
    """Load, assign, and save ``data/project_id_registry.csv``."""

    def __init__(self, path: Path, year_prefix: str):
        self.path = Path(path)
        self.year_prefix = str(year_prefix)
        self.rows: list[dict] = []
        self._by_key: dict[str, dict] = {}
        self._load()

    def _load(self) -> None:
        if not self.path.exists():
            return
        frame = pd.read_csv(self.path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        missing = [column for column in REGISTRY_COLUMNS if column not in frame.columns]
        if missing:
            raise ValueError(
                f"Registry {self.path} is missing columns: {', '.join(missing)}. "
                "Fix or remove the file before re-running so existing IDs are not renumbered."
            )
        for record in frame.to_dict(orient="records"):
            row = {column: str(record.get(column, "") or "") for column in REGISTRY_COLUMNS}
            self.rows.append(row)
            self._by_key[match_key(row["Supervisor_Email"], row["Project_Title"])] = row

    def resolve(self, email: str, title: str, response_id: str) -> tuple[str, bool]:
        """Reuse an ID for a known key, or allocate the next one.

        Returns ``(project_id, created_new)``.
        """
        key = match_key(email, title)
        existing = self._by_key.get(key)
        if existing:
            return existing["Project_ID"], False
        project_id = self._next_id()
        row = {
            "Project_ID": project_id,
            "Supervisor_Email": email or "",
            "Project_Title": title or "",
            "First_Seen_Response_Id": response_id or "",
            "Date_Assigned": date.today().isoformat(),
        }
        self.rows.append(row)
        self._by_key[key] = row
        return project_id, True

    def _next_id(self) -> str:
        prefix = f"SRP{self.year_prefix}-"
        numbers = []
        for row in self.rows:
            project_id = row["Project_ID"]
            if not project_id.startswith(prefix):
                continue
            suffix = project_id[len(prefix):]
            if suffix.isdigit():
                numbers.append(int(suffix))
        nxt = (max(numbers) if numbers else 0) + 1
        return f"{prefix}{nxt:03d}"

    def duplicate_project_ids(self) -> list[str]:
        counts = Counter(row["Project_ID"] for row in self.rows)
        return sorted(project_id for project_id, count in counts.items() if count > 1)

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.rows.sort(key=lambda row: row["Project_ID"])
        frame = pd.DataFrame(self.rows, columns=REGISTRY_COLUMNS)
        frame.to_csv(self.path, index=False, encoding="utf-8-sig")

    def export_lookup(self, path: Path, names_by_id: dict[str, str]) -> None:
        """Write the student/faculty reference sheet of current Project_IDs."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        records = []
        for row in sorted(self.rows, key=lambda item: item["Project_ID"]):
            records.append(
                {
                    "Project_ID": row["Project_ID"],
                    "Project_Title": row["Project_Title"],
                    "Supervisor_Name": names_by_id.get(row["Project_ID"], ""),
                    "Supervisor_Email": row["Supervisor_Email"],
                    "Date_Assigned": row["Date_Assigned"],
                }
            )
        frame = pd.DataFrame(
            records,
            columns=[
                "Project_ID",
                "Project_Title",
                "Supervisor_Name",
                "Supervisor_Email",
                "Date_Assigned",
            ],
        )
        frame.to_excel(path, index=False, sheet_name="Project_IDs")
