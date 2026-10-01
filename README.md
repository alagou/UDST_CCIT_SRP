# CCIT SRP Project-Status Reports

Daily command-line report for the CCIT Student Research Program. It reads Microsoft Forms Excel exports, assigns a permanent `Project_ID` to each faculty proposal, and writes an HTML dashboard, a master workbook, and a validation report.

## Setup

Python 3.11 or newer.

```bash
pip install -r requirements.txt
```

Copy the Faculty Proposal Form export into `input/`. The file name must match the glob in `config.py` (`*Proposal*Form*.xlsx`). The workbook shipped with this project is already in that folder when you first set the repo up from the 2026/2027 export.

## Daily run

```bash
python generate_srp_report.py
```

No arguments. Replace the Excel file in `input/` when a new Forms export arrives and run the command again. Outputs in `output/` are overwritten. Project IDs are not renumbered: `data/project_id_registry.csv` remembers each supervisor email + project title.

## Folder layout

```text
config.py                 settings, column maps, status labels and colors
cleaning.py               one cleaning function per field type
id_registry.py            SRP26-001 style IDs, saved between runs
merge.py                  combine forms, capacity, and status
validation.py             warnings and errors
html_dashboard.py         self-contained HTML dashboard
generate_srp_report.py    entry point
templates/dashboard.html.j2
input/                    drop Forms exports here
data/project_id_registry.csv
data/project_id_lookup.xlsx
output/SRP_Project_Status.html
output/SRP_Master_Data.xlsx
output/SRP_Validation_Report.xlsx
```

## What the run does

1. Find each configured form in `input/` by filename pattern.
2. Read columns by exact header name (`column_map` in `config.py`).
3. Clean text, emails, capacity, year/program, and student ID lists.
4. Reuse or assign `Project_ID` values (`SRP26-001`).
5. Start from the faculty name/ID list. When `*Student*Registration*.xlsx` is in `input/`, apply each row: Register adds that student to the selected project, and Unregister removes them. The latest submission wins. Until that file has responses, membership stays on the faculty list.
6. Apply an Admin Status Form label when that file exists (latest completion time wins).
7. Write the dashboard, master workbook, and validation report, and refresh `data/project_id_lookup.xlsx` so students can be given the right `Project_ID`.

Optional forms that are not on disk are skipped. The report still runs with only the faculty proposal file.

## Status labels

For each project, in this order:

1. Admin Form status, verbatim, recorded as `admin-override` in the validation report.
2. No students assigned: `Recruiting — No Students Assigned.`
3. Some seats left: `Recruiting.`
4. Assigned count at or above capacity: `Team Full.`
5. Students assigned but capacity is missing or not positive: `Action Needed`.

Green is recruiting, blue is team full, gray is closed or completed, and orange is action needed. The dashboard uses `status_style()` so the words and the color stay together.

## Adding another form later

Add a dictionary in `config.py` with `role`, `path`, `glob`, `sheet`, `optional`, `required_columns`, and `column_map`.

- `role: "projects"` feeds the project table (same shape as the faculty form).
- `role: "memberships"` feeds the membership table (same shape as the student interest form).
- `role: "registrations"` applies the supervisor register/unregister form.
- `role: "status_overrides"` feeds admin status.

The reader loops over every form dictionary. You do not edit the loader to add a file of one of these roles.

## SharePoint / OneDrive later

Set `input_folder` in `config.py` to the synced folder. Discovery uses globs such as `*Proposal*Form*.xlsx`, so a new export name still matches. `fetch_latest_forms()` in `generate_srp_report.py` is the stub where a Microsoft Graph download can be added. This version does not call Graph.

## Libraries

- **pandas + openpyxl** read and write Excel by column name, so a reordered form still lines up.
- **jinja2** renders the HTML dashboard from `templates/dashboard.html.j2`. The page inlines its CSS and JavaScript, so opening the file in a browser is enough.
