"""Combine Faculty, Student, and Admin forms on Project_ID.

Faculty rows become projects. Student rows become memberships. Until a
Student Form exists for a project, the faculty free-text name/ID list is
the fallback and is marked ``faculty_fallback``. Admin status, when
present, overrides the computed recruitment label (latest timestamp wins).
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

from cleaning import (
    clean_email,
    clean_multiline,
    clean_text,
    normalize_project_id,
    normalize_student_id,
    pair_names_ids,
    parse_capacity,
    parse_registration_action,
    split_multi_value,
    split_year_program,
)
from config import status_style
from id_registry import ProjectRegistry, match_key
from validation import add_issue

PROJECT_COLUMNS = [
    "Project_ID",
    "Title",
    "Supervisor",
    "Supervisor_Email",
    "Co_Supervisor",
    "Status_Label",
    "Status_Source",
    "Max_Students",
    "Assigned_Count",
    "Available_Positions",
    "Preferred_Program",
    "Preferred_Year",
    "Required_Skills",
    "Last_Updated",
    "Source_Response_Id",
    "Membership_Source",
]
MEMBER_COLUMNS = [
    "Project_ID",
    "Student_ID",
    "Student_Name",
    "Membership_Status",
    "Source_Form",
]


def build_dataset(faculty_df, student_df, admin_df, registry: ProjectRegistry, config: dict, registration_df=None):
    """Return ``(projects, memberships, issues)``."""
    issues: list = []
    records = _faculty_records(faculty_df) if faculty_df is not None else []
    winners, duplicate_notes = _collapse_duplicates(records)
    winners.sort(key=lambda record: (record["completion_time"], record["response_id"]))

    projects = []
    created = reused = 0
    for record in winners:
        project_id, is_new = registry.resolve(record["email"] or "", record["title"] or "", record["response_id"])
        created += int(is_new)
        reused += int(not is_new)
        record["project_id"] = project_id
        projects.append(_project_shell(record))
    id_by_key = {project["match_key"]: project["project_id"] for project in projects}
    for key, raw in duplicate_notes:
        add_issue(
            issues,
            "INFO",
            id_by_key.get(key, ""),
            "Completion time",
            "Duplicate faculty submission treated as an update",
            raw,
        )

    project_ids = {project["project_id"] for project in projects}
    student_rows, projects_with_students = _student_memberships(
        student_df, project_ids, issues, config["student_id_length"]
    )
    students_by_project: dict[str, list] = {}
    for row in student_rows:
        students_by_project.setdefault(row["project_id"], []).append(row)

    memberships = []
    for project in projects:
        if project["project_id"] in projects_with_students:
            members = students_by_project.get(project["project_id"], [])
            project["membership_source"] = "student_form"
            _flag_membership_conflict(project, members, issues, config["student_id_length"])
        else:
            members = _faculty_members(project, config["student_id_length"])
            project["membership_source"] = "faculty_fallback"
        project["members"] = members
        memberships.extend(members)

    _apply_registrations(
        projects, registration_df, project_ids, issues, config["student_id_length"]
    )
    memberships = [member for project in projects for member in project["members"]]

    overrides = _admin_overrides(admin_df, project_ids, issues)
    for project in projects:
        _apply_status(project, overrides.get(project["project_id"]), config["status_labels"])

    projects.sort(key=lambda project: project["project_id"])
    return projects, memberships, issues, created, reused


def _faculty_records(frame) -> list[dict]:
    records = []
    for row in frame.to_dict(orient="records"):
        email, email_warning = clean_email(row.get("supervisor_email"))
        title, _ = clean_text(row.get("title"))
        name, _ = clean_text(row.get("supervisor_name"))
        maximum, capacity_warning = parse_capacity(row.get("capacity"))
        (years, programs), year_warning = split_year_program(row.get("year_program"))
        description, _ = clean_multiline(row.get("description"))
        objectives, _ = clean_multiline(row.get("objectives"))
        skills, _ = clean_multiline(row.get("skills"))
        comments, _ = clean_multiline(row.get("comments"))
        co_supervisor, _ = clean_text(row.get("co_supervisor"))
        response_id, _ = clean_text(row.get("response_id"))
        completion, _ = clean_text(row.get("completion_time"))
        records.append(
            {
                "response_id": response_id or "",
                "completion_time": completion or "",
                "email": email or "",
                "email_warning": email_warning,
                "title": title or "",
                "name": name or "",
                "maximum": maximum,
                "capacity_warning": capacity_warning,
                "years": years,
                "programs": programs,
                "year_warning": year_warning,
                "description": description,
                "objectives": objectives,
                "skills": skills,
                "comments": comments,
                "co_supervisor": co_supervisor,
                "raw_email": row.get("supervisor_email") or "",
                "raw_title": row.get("title") or "",
                "raw_name": row.get("supervisor_name") or "",
                "raw_capacity": row.get("capacity") or "",
                "raw_year_program": row.get("year_program") or "",
                "raw_student_names": row.get("student_names") or "",
                "raw_student_ids": row.get("student_ids") or "",
                "match_key": match_key(email, title),
            }
        )
    return records


def _collapse_duplicates(records: list[dict]):
    """Keep the latest Completion time for each faculty match key."""
    groups: dict[str, list] = {}
    for record in records:
        groups.setdefault(record["match_key"], []).append(record)
    winners = []
    notes = []
    for key, group in groups.items():
        group.sort(key=lambda record: (record["completion_time"], record["response_id"]))
        winner = group[-1]
        winners.append(winner)
        for older in group[:-1]:
            notes.append(
                (
                    key,
                    f"kept response {winner['response_id']} at {winner['completion_time']}; "
                    f"superseded response {older['response_id']} at {older['completion_time']}",
                )
            )
    return winners, notes


def _project_shell(record: dict) -> dict:
    return {
        "project_id": record["project_id"],
        "match_key": record["match_key"],
        "title": record["title"],
        "supervisor_name": record["name"],
        "supervisor_email": record["email"],
        "email_warning": record["email_warning"],
        "co_supervisor": record["co_supervisor"],
        "description": record["description"],
        "objectives": record["objectives"],
        "skills": record["skills"],
        "years": record["years"],
        "programs": record["programs"],
        "year_warning": record["year_warning"],
        "max_students": record["maximum"],
        "capacity_warning": record["capacity_warning"],
        "comments": record["comments"],
        "completion_time": record["completion_time"],
        "response_id": record["response_id"],
        "raw_supervisor_name": record["raw_name"],
        "raw_supervisor_email": record["raw_email"],
        "raw_title": record["raw_title"],
        "raw_capacity": record["raw_capacity"],
        "raw_year_program": record["raw_year_program"],
        "raw_student_names": record["raw_student_names"],
        "raw_student_ids": record["raw_student_ids"],
        "name_id_mismatch": None,
        "members": [],
        "membership_source": "faculty_fallback",
    }


def _student_memberships(frame, project_ids: set, issues: list, id_length: int):
    rows = []
    covered = set()
    if frame is None:
        return rows, covered
    for row in frame.to_dict(orient="records"):
        project_id, format_warning = normalize_project_id(row.get("project_id"))
        response_id, _ = clean_text(row.get("response_id"))
        if not project_id:
            add_issue(issues, "ERROR", "", "Project_ID", "Missing Project_ID", response_id or "")
            continue
        if format_warning:
            add_issue(issues, "ERROR", project_id, "Project_ID", "Malformed Project_ID", row.get("project_id"))
            continue
        if project_id not in project_ids:
            add_issue(issues, "ERROR", project_id, "Project_ID", "Unknown Project_ID", row.get("project_id"))
            continue
        name, _ = clean_text(row.get("student_name"))
        student_id, id_warning = normalize_student_id(row.get("student_id"), id_length)
        rows.append(
            _member(
                project_id,
                name,
                student_id,
                id_warning is None and bool(student_id),
                "student_form",
                False,
                row.get("student_id"),
            )
        )
        covered.add(project_id)
    return rows, covered


def _faculty_members(project: dict, id_length: int) -> list[dict]:
    names, _ = split_multi_value(project.get("raw_student_names"))
    raw_ids, _ = split_multi_value(project.get("raw_student_ids"))
    pairs, mismatch = pair_names_ids(names, raw_ids)
    if mismatch:
        project["name_id_mismatch"] = mismatch
    members = []
    for pair in pairs:
        student_id, id_warning = normalize_student_id(pair["student_id"], id_length)
        valid = id_warning is None and bool(student_id) and not pair["incomplete"]
        # A valid ID still counts toward capacity even when names could not be aligned.
        if id_warning is None and student_id and pair["incomplete"]:
            valid = True
        members.append(
            _member(
                project["project_id"],
                pair["name"],
                student_id,
                valid,
                "faculty_form",
                pair["incomplete"] or bool(id_warning),
                pair["student_id"],
            )
        )
    return members


def _member(project_id, name, student_id, id_valid, source, incomplete, raw_student_id) -> dict:
    status = "Assigned" if id_valid and not incomplete else "Incomplete"
    return {
        "project_id": project_id,
        "student_name": name,
        "student_id": student_id,
        "id_valid": bool(id_valid),
        "membership_status": status,
        "source_form": source,
        "raw_student_id": raw_student_id or "",
    }


def _flag_membership_conflict(project: dict, members: list, issues: list, id_length: int) -> None:
    faculty_ids = set()
    for part in split_multi_value(project.get("raw_student_ids"))[0]:
        student_id, warning = normalize_student_id(part, id_length)
        if student_id and warning is None:
            faculty_ids.add(student_id)
    student_ids = {member["student_id"] for member in members if member.get("id_valid")}
    if faculty_ids and faculty_ids != student_ids:
        add_issue(
            issues,
            "WARNING",
            project["project_id"],
            "Student ID(s)",
            "Faculty member list differs from Student Form; Student Form kept",
            f"faculty={sorted(faculty_ids)}; student_form={sorted(student_ids)}",
        )


def _apply_registrations(projects: list, frame, project_ids: set, issues: list, id_length: int) -> None:
    """Apply supervisor register/unregister rows on top of the current roster.

    The faculty list remains the starting roster. Each registration row changes
    one student. If the same student is submitted more than once, the latest
    Completion time wins.
    """
    events = _registration_events(frame, project_ids, issues, id_length)
    if not events:
        return
    by_project: dict[str, list] = {}
    for event in events:
        by_project.setdefault(event["project_id"], []).append(event)
    for project in projects:
        project_events = by_project.get(project["project_id"])
        if not project_events:
            continue
        _apply_project_events(project, project_events, issues)
        project["membership_source"] = "registration_form"


def _registration_events(frame, project_ids: set, issues: list, id_length: int) -> list:
    if frame is None or frame.empty:
        return []
    events = []
    for row in frame.to_dict(orient="records"):
        action, action_warning = parse_registration_action(row.get("action"))
        project_id, format_warning = normalize_project_id(row.get("project_id"))
        response_id, _ = clean_text(row.get("response_id"))
        completion, _ = clean_text(row.get("completion_time"))
        if action_warning:
            add_issue(issues, "ERROR", project_id or "", "Action", "Missing or unmapped registration action", row.get("action"))
            continue
        if not project_id:
            add_issue(issues, "ERROR", "", "Project", "Missing Project_ID", response_id or "")
            continue
        if format_warning:
            add_issue(issues, "ERROR", "", "Project", "Malformed Project_ID", row.get("project_id"))
            continue
        if project_id not in project_ids:
            add_issue(issues, "ERROR", project_id, "Project", "Unknown Project_ID", row.get("project_id"))
            continue
        name, _ = clean_text(row.get("student_name"))
        student_id, id_warning = normalize_student_id(row.get("student_id"), id_length)
        if not student_id:
            add_issue(issues, "ERROR", project_id, "Student ID", "Missing Student_ID", name or response_id or "")
            continue
        if id_warning:
            add_issue(issues, "WARNING", project_id, "Student ID", "Invalid Student_ID", row.get("student_id"))
        events.append(
            {
                "action": action,
                "project_id": project_id,
                "student_name": name,
                "student_id": student_id,
                "id_valid": id_warning is None,
                "completion_time": completion or "",
                "response_id": response_id or "",
                "raw_student_id": row.get("student_id"),
            }
        )
    return events


def _apply_project_events(project: dict, events: list, issues: list) -> None:
    grouped: dict[str, list] = {}
    for event in events:
        grouped.setdefault(event["student_id"], []).append(event)
    for student_id, group in grouped.items():
        group.sort(key=lambda event: (event["completion_time"], event["response_id"]))
        for older in group[:-1]:
            add_issue(
                issues,
                "INFO",
                project["project_id"],
                "Action",
                "Superseded registration; latest Completion time kept",
                f"{older['action']} {student_id} at {older['completion_time']}",
            )
        _apply_one_event(project, group[-1], issues)


def _apply_one_event(project: dict, event: dict, issues: list) -> None:
    member = next(
        (
            item
            for item in project["members"]
            if item.get("student_id") == event["student_id"] and item.get("membership_status") != "Unregistered"
        ),
        None,
    )
    if event["action"] == "unregister":
        if member is None:
            add_issue(
                issues,
                "WARNING",
                project["project_id"],
                "Student ID",
                "Unregister requested for a student who is not assigned to this project",
                event["raw_student_id"],
            )
            return
        member["membership_status"] = "Unregistered"
        member["source_form"] = "registration_form"
        if event["student_name"]:
            member["student_name"] = event["student_name"]
        return
    if member is None:
        project["members"].append(
            _member(
                project["project_id"],
                event["student_name"],
                event["student_id"],
                event["id_valid"],
                "registration_form",
                not event["id_valid"],
                event["raw_student_id"],
            )
        )
        return
    member["membership_status"] = "Assigned" if event["id_valid"] else "Incomplete"
    member["id_valid"] = event["id_valid"]
    member["source_form"] = "registration_form"
    if event["student_name"]:
        member["student_name"] = event["student_name"]


def _admin_overrides(frame, project_ids: set, issues: list) -> dict[str, str]:
    """Latest admin status per Project_ID. Earlier rows are superseded."""
    if frame is None:
        return {}
    grouped: dict[str, list] = {}
    for row in frame.to_dict(orient="records"):
        project_id, format_warning = normalize_project_id(row.get("project_id"))
        if not project_id:
            add_issue(issues, "ERROR", "", "Project_ID", "Missing Project_ID", row.get("response_id") or "")
            continue
        if format_warning:
            add_issue(issues, "ERROR", project_id, "Project_ID", "Malformed Project_ID", row.get("project_id"))
            continue
        if project_id not in project_ids:
            add_issue(issues, "ERROR", project_id, "Project_ID", "Unknown Project_ID", row.get("project_id"))
            continue
        status, _ = clean_text(row.get("status"))
        if not status:
            add_issue(issues, "WARNING", project_id, "Status", "Admin status is blank", row.get("response_id") or "")
            continue
        completion, _ = clean_text(row.get("completion_time"))
        response_id, _ = clean_text(row.get("response_id"))
        grouped.setdefault(project_id, []).append((completion or "", response_id or "", status))

    overrides = {}
    for project_id, updates in grouped.items():
        updates.sort()
        for older in updates[:-1]:
            add_issue(
                issues,
                "INFO",
                project_id,
                "Status",
                "Superseded admin update",
                f"response {older[1]} at {older[0]} status {older[2]}",
            )
        overrides[project_id] = updates[-1][2]
    return overrides


def _apply_status(project: dict, admin_status: str | None, labels: dict) -> None:
    """Set the displayed recruitment status.

    Precedence, also written to the validation report when an override applies:
    1. An Admin Form status for this Project_ID is used verbatim and the
       computed label is skipped. The computed label is still stored so the
       conflict can be audited.
    2. Otherwise if nobody is assigned: Recruiting — No Students Assigned.
    3. Otherwise if seats remain: Recruiting.
    4. Otherwise if the team is at or over capacity: Team Full.
    5. If capacity is missing or not positive and students are already
       assigned, none of the tests above match, so the label is Action Needed.
    """
    valid_ids = {
        member["student_id"]
        for member in project["members"]
        if member.get("id_valid") and member.get("membership_status") == "Assigned"
    }
    assigned = len(valid_ids)
    maximum = project["max_students"]
    project["assigned_count"] = assigned
    if maximum is None:
        project["available_positions"] = None
    else:
        project["available_positions"] = max(maximum - assigned, 0)
    project["is_empty"] = assigned == 0
    project["is_full"] = maximum is not None and maximum > 0 and assigned >= maximum
    project["is_partial"] = maximum is not None and maximum > 0 and 0 < assigned < maximum

    if project["is_empty"]:
        computed = labels["empty"]
    elif project["is_partial"]:
        computed = labels["partial"]
    elif project["is_full"]:
        computed = labels["full"]
    else:
        computed = labels["action_needed"]
    project["computed_label"] = computed
    if admin_status:
        project["status_label"] = admin_status
        project["status_source"] = "admin-override"
    else:
        project["status_label"] = computed
        project["status_source"] = "computed"


def build_view(projects: list, issues: list):
    """Build card dicts, summary counts, and the program filter list."""
    known = {project["project_id"] for project in projects}
    attention = set()
    for issue in issues:
        if issue["Severity"] not in {"WARNING", "ERROR"}:
            continue
        for project_id in str(issue.get("Project_ID") or "").split(","):
            project_id = project_id.strip()
            if project_id in known:
                attention.add(project_id)

    cards = [_card(project) for project in projects]
    summary = {
        "total_projects": len(projects),
        "total_students_assigned": sum(project["assigned_count"] for project in projects),
        "projects_recruiting": sum(1 for project in projects if "recruit" in project["status_label"].lower()),
        "projects_full": sum(1 for project in projects if "team full" in project["status_label"].lower()),
        "projects_with_no_students": sum(1 for project in projects if project["is_empty"]),
        "total_available_positions": sum(project["available_positions"] or 0 for project in projects),
        "projects_requiring_attention": len(attention),
    }
    programs = []
    seen = set()
    for card in cards:
        for program in card["programs"]:
            key = program.lower()
            if key not in seen:
                seen.add(key)
                programs.append(program)
    return cards, summary, programs


def _card(project: dict) -> dict:
    color, label = status_style(project["status_label"])
    members = []
    for member in project["members"]:
        if member.get("membership_status") == "Unregistered":
            continue
        name = member["student_name"] or "Name missing"
        student_id = member["student_id"] or "ID missing"
        members.append({"name": name, "student_id": student_id})
    programs = project.get("programs") or []
    years = project.get("years") or []
    maximum = project["max_students"]
    blank = "\u2014"
    program_text = ", ".join(programs) if programs else blank
    year_text = ", ".join(years) if years else blank
    return {
        "project_id": project["project_id"],
        "title": project["title"] or "Untitled project",
        "supervisor_name": project["supervisor_name"] or "Supervisor name missing",
        "supervisor_email": project["supervisor_email"] or "email missing",
        "co_supervisor": project.get("co_supervisor"),
        "status_label": label,
        "status_color": color,
        "assigned": project["assigned_count"],
        "maximum_display": blank if maximum is None else str(maximum),
        "available_display": blank if maximum is None else str(project["available_positions"]),
        "members": members,
        "skills": project.get("skills") or "",
        "program_year": f"{program_text} / {year_text}",
        "programs": programs,
        "programs_key": "|".join(program.lower() for program in programs),
        "description": project.get("description") or "",
        "objectives": project.get("objectives") or "",
        "comments": project.get("comments") or "",
        "membership_source": project["membership_source"],
        "bucket": _bucket(label),
        "is_empty": project["is_empty"],
        "student_search": " ".join(
            f"{member['student_name'] or ''} {member['student_id'] or ''}" for member in project["members"]
        ).lower(),
    }


def _bucket(label: str) -> str:
    low = (label or "").lower()
    if any(word in low for word in ("closed", "completed", "cancelled", "canceled")):
        return "closed"
    if "team full" in low:
        return "full"
    if "recruit" in low:
        return "recruiting"
    if "action needed" in low or "incomplete" in low or "hold" in low:
        return "action"
    return "other"


def write_master(path: Path, projects: list, memberships: list) -> None:
    """Write the Projects and Memberships sheets joined by Project_ID."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    project_rows = []
    for project in projects:
        maximum = project["max_students"]
        available = project["available_positions"]
        project_rows.append(
            {
                "Project_ID": project["project_id"],
                "Title": project["title"],
                "Supervisor": project["supervisor_name"],
                "Supervisor_Email": project["supervisor_email"],
                "Co_Supervisor": project.get("co_supervisor") or "",
                "Status_Label": project["status_label"],
                "Status_Source": project["status_source"],
                "Max_Students": "" if maximum is None else maximum,
                "Assigned_Count": project["assigned_count"],
                "Available_Positions": "" if available is None else available,
                "Preferred_Program": ", ".join(project.get("programs") or []),
                "Preferred_Year": ", ".join(project.get("years") or []),
                "Required_Skills": project.get("skills") or "",
                "Last_Updated": project.get("completion_time") or "",
                "Source_Response_Id": project.get("response_id") or "",
                "Membership_Source": project["membership_source"],
            }
        )
    member_rows = [
        {
            "Project_ID": member["project_id"],
            "Student_ID": member["student_id"] or "",
            "Student_Name": member["student_name"] or "",
            "Membership_Status": member["membership_status"],
            "Source_Form": member["source_form"],
        }
        for member in memberships
    ]
    projects_frame = pd.DataFrame(project_rows, columns=PROJECT_COLUMNS)
    members_frame = pd.DataFrame(member_rows, columns=MEMBER_COLUMNS)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        projects_frame.to_excel(writer, sheet_name="Projects", index=False)
        members_frame.to_excel(writer, sheet_name="Memberships", index=False)
        _style(writer.sheets["Projects"])
        _style(writer.sheets["Memberships"])


def _style(worksheet) -> None:
    header_fill = PatternFill("solid", fgColor="0F2744")
    header_font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for column in worksheet.columns:
        letter = column[0].column_letter
        worksheet.column_dimensions[letter].width = 22 if letter != "A" else 16
    worksheet.column_dimensions["B"].width = 42
    worksheet.column_dimensions["M"].width = 40
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
