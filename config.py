"""Configuration for the CCIT SRP project-status pipeline.

Business rules, file locations, and form column mappings live here.
The rest of the pipeline reads this module and does not hard-code
filenames, column headers, or status labels.

Add another Microsoft Forms export by appending a new dictionary with a
``role`` of ``projects``, ``memberships``, or ``status_overrides``. Set
``share_url`` to the workbook's SharePoint sharing link and the report
downloads it on each run. Discovery still uses each form's ``glob``.
"""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent

# Exact Forms headers. The email header contains an en dash (U+2013)
# and a non-breaking space (U+00A0). Matching is by this full string.
SUPERVISOR_NAME = "Faculty Supervisor \u2013 Full Name"
SUPERVISOR_EMAIL = "Faculty Supervisor \u2013\u00a0UDST Email Address"

CONFIG = {
    "input_folder": "input",
    "academic_year": "26",
    "membership_exclusive": True,
    "student_id_length": 8,
    "output_folder": "output",
    "registry_path": "data/project_id_registry.csv",
    "lookup_path": "data/project_id_lookup.xlsx",
    "outputs": {
        "html": "SRP_Project_Status.html",
        "master": "SRP_Master_Data.xlsx",
        "validation": "SRP_Validation_Report.xlsx",
    },
    # Displayed labels. Precedence is documented in merge._apply_status.
    "status_labels": {
        "empty": "Recruiting \u2014 No Students Assigned.",
        "partial": "Recruiting.",
        "full": "Team Full.",
        "action_needed": "Action Needed",
    },
    # Color is never the only signal; status_style always returns the label too.
    "status_colors": {
        "recruiting": "#15803D",
        "action_needed": "#D97706",
        "full": "#1D4ED8",
        "closed": "#6B7280",
    },
    "faculty_form": {
        "role": "projects",
        "path": "input/CCIT-SRP-Faculty-Proposal-Form.xlsx",
        "glob": "*Proposal*Form*.xlsx",
        "share_url": (
            "https://myudst-my.sharepoint.com/:x:/g/personal/60107348_udst_edu_qa/"
            "IQCvmB0oEGLdTKIRpYela4k2AQtETJ8c1OC9VL6eL9xItO4"
        ),
        "sheet": "Sheet1",
        "optional": False,
        "required_columns": ["Project Title", SUPERVISOR_NAME, SUPERVISOR_EMAIL],
        "optional_columns": [
            "Co-Supervisor(s)",
            "Additional Comments",
            "Project Description",
            "Research Objectives & Expected Outcomes",
            "Required Skills / Knowledge Areas",
            "Preferred Student Year & Program",
            "Maximum Number of Students",
            "Selected/Preferred Student Name(s)",
            "Student ID(s)",
            "Student Recruitment Status",
            "Project Status",
        ],
        "column_map": {
            "Id": "response_id",
            "Start time": "start_time",
            "Completion time": "completion_time",
            "Email": "submitter_email",
            "Name": "submitter_name",
            SUPERVISOR_NAME: "supervisor_name",
            SUPERVISOR_EMAIL: "supervisor_email",
            "Project Title": "title",
            "Project Description": "description",
            "Research Objectives & Expected Outcomes": "objectives",
            "Co-Supervisor(s)": "co_supervisor",
            "Project Status": "faculty_status",
            "Required Skills / Knowledge Areas": "skills",
            "Preferred Student Year & Program": "year_program",
            "Maximum Number of Students": "capacity",
            "Student Recruitment Status": "recruitment_status",
            "Selected/Preferred Student Name(s)": "student_names",
            "Student ID(s)": "student_ids",
            "Additional Comments": "comments",
            "SRP Project Eligibility Confirmation": "eligibility",
        },
    },
    "student_form": {
        "role": "memberships",
        "path": "input/CCIT-SRP-Student-Interest-Form.xlsx",
        "glob": "*Student*Interest*Form*.xlsx",
        "sheet": "Sheet1",
        "optional": True,
        "required_columns": ["Project_ID", "Student Name", "Student ID"],
        "optional_columns": ["Completion time", "Id"],
        "column_map": {
            "Id": "response_id",
            "Completion time": "completion_time",
            "Project_ID": "project_id",
            "Student Name": "student_name",
            "Student ID": "student_id",
        },
    },
    "registration_form": {
        "role": "registrations",
        "path": "input/CCIT-SRP-Student-Registration-Form.xlsx",
        "glob": "*Student*Registration*.xlsx",
        "sheet": "Sheet1",
        "optional": True,
        "required_columns": ["Action", "Project", "Student ID", "Student full name"],
        "optional_columns": ["Id", "Start time", "Completion time", "Email", "Name"],
        "column_map": {
            "Id": "response_id",
            "Start time": "start_time",
            "Completion time": "completion_time",
            "Email": "submitter_email",
            "Name": "submitter_name",
            "Action": "action",
            "Project": "project_id",
            "Student ID": "student_id",
            "Student full name": "student_name",
        },
    },
    "admin_form": {
        "role": "status_overrides",
        "path": "input/CCIT-SRP-Admin-Status-Form.xlsx",
        "glob": "*Admin*Status*Form*.xlsx",
        "sheet": "Sheet1",
        "optional": True,
        "required_columns": ["Project_ID", "Status"],
        "optional_columns": ["Completion time", "Id", "Comments"],
        "column_map": {
            "Id": "response_id",
            "Completion time": "completion_time",
            "Project_ID": "project_id",
            "Status": "status",
            "Comments": "comments",
        },
    },
}


def status_style(label: str) -> tuple[str, str]:
    """Return ``(color_hex, text_label)`` for a status label.

    The dashboard uses this helper so a status is always a color plus the
    same words. Unknown labels fall back to the action-needed color and
    still keep their original text.
    """
    text = (label or "").strip() or CONFIG["status_labels"]["action_needed"]
    low = text.lower()
    colors = CONFIG["status_colors"]
    if any(word in low for word in ("action needed", "incomplete")):
        return colors["action_needed"], text
    if any(word in low for word in ("closed", "completed", "cancelled", "canceled")):
        return colors["closed"], text
    if "hold" in low:
        return colors["action_needed"], text
    if "team full" in low or low == "full":
        return colors["full"], text
    if "recruit" in low or "active" in low:
        return colors["recruiting"], text
    return colors["action_needed"], text


def iter_form_configs(config: dict | None = None):
    """Yield ``(name, form_config)`` for every configured form."""
    source = CONFIG if config is None else config
    for name, value in source.items():
        if isinstance(value, dict) and "column_map" in value and "role" in value:
            yield name, value


def project_path(relative: str) -> Path:
    """Resolve a config path relative to the project root."""
    path = Path(relative)
    if path.is_absolute():
        return path
    return ROOT / path
