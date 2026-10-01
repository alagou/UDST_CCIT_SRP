"""Self-contained HTML dashboard for the SRP project cards."""

from __future__ import annotations

import base64
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape

ROOT = Path(__file__).resolve().parent
TEMPLATE_DIR = ROOT / "templates"
LOGO_PATH = ROOT / "assets" / "udst-logo.svg"


def build_html(path: Path, cards: list, summary: dict, programs: list, generated_at: str) -> None:
    """Render ``SRP_Project_Status.html`` with inline CSS and JavaScript."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    environment = Environment(
        loader=FileSystemLoader(TEMPLATE_DIR),
        autoescape=select_autoescape(["html", "j2"]),
    )
    html = environment.get_template("dashboard.html.j2").render(
        cards=cards,
        summary=summary,
        programs=programs,
        generated_at=generated_at,
        logo_src=_logo_src(),
        contact_email="ala.gouissem@udst.edu.qa",
    )
    path.write_text(html, encoding="utf-8")


def _logo_src() -> str:
    encoded = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
    return f"data:image/svg+xml;base64,{encoded}"
