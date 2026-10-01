"""Web entrypoint for Vercel. Serves the generated SRP dashboard."""

from __future__ import annotations

import os
import traceback
from pathlib import Path

from flask import Flask, Response

from config import CONFIG, project_path
from generate_srp_report import main as generate_report

app = Flask(__name__)

if os.environ.get("VERCEL"):
    scratch = Path("/tmp/srp")
    CONFIG["input_folder"] = str(scratch / "input")
    CONFIG["output_folder"] = str(scratch / "output")
    CONFIG["registry_path"] = str(scratch / "data" / "project_id_registry.csv")
    CONFIG["lookup_path"] = str(scratch / "data" / "project_id_lookup.xlsx")

_cached_html: str | None = None


def _dashboard() -> str:
    global _cached_html
    if _cached_html is None:
        generate_report()
        html_path = project_path(CONFIG["output_folder"]) / CONFIG["outputs"]["html"]
        _cached_html = html_path.read_text(encoding="utf-8")
    return _cached_html


@app.get("/")
def home():
    try:
        return Response(_dashboard(), mimetype="text/html")
    except Exception:
        return Response("<pre>" + traceback.format_exc() + "</pre>", status=500)
