"""Build a TPM weekly status report from an Excel workbook.

Reads the Summary sheet and the Risk and Issue sheet, then writes one
detailed text section per project to Weekly_Report.txt. Each project
lists client, status, and next milestone, then issues, risks, and
recommended actions.
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# Sheet names in TPM.xlsx. Matching ignores extra whitespace.
SUMMARY_SHEET = "Summary"
RISK_SHEET = "Risk and Issue"

# Canonical column names after headers are stripped and lower-cased.
SUMMARY_COLUMNS = {
    "project": "Project",
    "client": "Client",
    "status": "Status",
    "next milestone": "Next Milestone",
    "date": "Date",
    "main problem": "Main Problem",
}
RISK_COLUMNS = {
    "id": "ID",
    "project": "Project",
    "type": "Type",
    "description": "Description",
    "probability": "Probability",
    "impact": "Impact",
    "owner": "Owner",
    "action": "Action",
    "due date": "Due Date",
}

# Rows whose Type is neither Issue nor Risk are still reported, grouped
# under the type name (for example "Change Request").
ISSUE_TYPE = "issue"
RISK_TYPE = "risk"


def parse_args() -> argparse.Namespace:
    """Accept optional workbook and output paths. Defaults sit beside this script."""
    script_dir = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(
        description="Generate a TPM weekly report from Summary and Risk and Issue sheets."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=script_dir / "TPM.xlsx",
        help="Path to the TPM workbook (default: TPM.xlsx next to this script).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=script_dir / "Weekly_Report.txt",
        help="Path of the text report to write (default: Weekly_Report.txt).",
    )
    return parser.parse_args()


def load_sheet(workbook: Path, sheet_name: str) -> pd.DataFrame:
    """Load one worksheet with pandas and openpyxl, then drop fully blank rows."""
    if not workbook.is_file():
        raise FileNotFoundError(f"Workbook not found: {workbook}")

    try:
        frame = pd.read_excel(workbook, sheet_name=sheet_name, engine="openpyxl")
    except ValueError as exc:
        # pandas raises ValueError when the requested sheet is missing.
        raise ValueError(
            f"Sheet '{sheet_name}' was not found in {workbook.name}."
        ) from exc

    frame = frame.dropna(how="all")
    frame.columns = [str(column).strip() for column in frame.columns]
    return frame


def rename_columns(frame: pd.DataFrame, mapping: dict[str, str], sheet_name: str) -> pd.DataFrame:
    """Rename headers to a stable set and fail if a required column is absent."""
    lookup = {column.lower(): column for column in frame.columns}
    missing = [canonical for key, canonical in mapping.items() if key not in lookup]
    if missing:
        raise ValueError(
            f"Sheet '{sheet_name}' is missing columns: {', '.join(missing)}."
        )

    renamed = frame.rename(columns={lookup[key]: canonical for key, canonical in mapping.items()})
    return renamed.loc[:, list(mapping.values())].copy()


def cell_text(value: object) -> str:
    """Turn a cell into display text. Blank cells and NaN become an empty string."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%d %b %Y")
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def project_key(name: object) -> str:
    """Key used to join Summary rows with Risk and Issue rows."""
    return cell_text(name).casefold()


def format_milestone(milestone: str, milestone_date: str) -> str:
    """Show the next milestone with its date when the Summary sheet has one."""
    if milestone and milestone_date:
        return f"{milestone} ({milestone_date})"
    return milestone or milestone_date or "Not recorded"


def field_line(label: str, value: str, width: int) -> str:
    """One labelled field, padded so values in a block line up."""
    return f"  {label:<{width}} : {value}"


def issue_lines(row: pd.Series) -> list[str]:
    """Format one issue with id, description, owner, impact, action, and due date."""
    width = len("Description")
    return [
        field_line("Issue ID", cell_text(row["ID"]) or "—", width),
        field_line("Description", cell_text(row["Description"]) or "No description", width),
        field_line("Owner", cell_text(row["Owner"]) or "—", width),
        field_line("Impact", cell_text(row["Impact"]) or "—", width),
        field_line("Action", cell_text(row["Action"]) or "—", width),
        field_line("Due Date", cell_text(row["Due Date"]) or "—", width),
    ]


def risk_lines(row: pd.Series) -> list[str]:
    """Format one risk, including probability and the mitigation action."""
    width = len("Mitigation Action")
    return [
        field_line("Risk ID", cell_text(row["ID"]) or "—", width),
        field_line("Description", cell_text(row["Description"]) or "No description", width),
        field_line("Probability", cell_text(row["Probability"]) or "—", width),
        field_line("Impact", cell_text(row["Impact"]) or "—", width),
        field_line("Owner", cell_text(row["Owner"]) or "—", width),
        field_line("Mitigation Action", cell_text(row["Action"]) or "—", width),
        field_line("Due Date", cell_text(row["Due Date"]) or "—", width),
    ]


def other_lines(row: pd.Series) -> list[str]:
    """Format a log row that is neither an issue nor a risk."""
    width = len("Description")
    return [
        field_line("ID", cell_text(row["ID"]) or "—", width),
        field_line("Description", cell_text(row["Description"]) or "No description", width),
        field_line("Probability", cell_text(row["Probability"]) or "—", width),
        field_line("Impact", cell_text(row["Impact"]) or "—", width),
        field_line("Owner", cell_text(row["Owner"]) or "—", width),
        field_line("Action", cell_text(row["Action"]) or "—", width),
        field_line("Due Date", cell_text(row["Due Date"]) or "—", width),
    ]


def recommended_action_line(row: pd.Series) -> str | None:
    """Return one recommended-action line, or None when the Action cell is empty."""
    action = cell_text(row["Action"])
    if not action:
        return None
    item_id = cell_text(row["ID"]) or "—"
    kind = cell_text(row["Type"]) or "Item"
    owner = cell_text(row["Owner"]) or "Unassigned"
    due = cell_text(row["Due Date"]) or "No due date"
    return f"  - [{item_id}] {action} ({kind}) | Owner: {owner} | Due Date: {due}"


def section_block(title: str, rows: pd.DataFrame, formatter) -> list[str]:
    """Render a labelled list. An empty frame becomes a single 'None' line."""
    lines = [title]
    if rows.empty:
        lines.append("  None")
        return lines
    for index, (_, row) in enumerate(rows.iterrows()):
        if index:
            lines.append("")
        lines.extend(formatter(row))
    return lines


def rows_for_project(log: pd.DataFrame, key: str) -> pd.DataFrame:
    """Filter the risk log to one project, preserving sheet order."""
    if log.empty:
        return log
    mask = log["Project"].map(project_key) == key
    return log.loc[mask]


def build_project_section(summary_row: pd.Series | None, log_rows: pd.DataFrame, fallback_name: str) -> list[str]:
    """Assemble the weekly block for a single project."""
    if summary_row is None:
        project_name = fallback_name
        client = "Not recorded"
        status = "Not recorded"
        milestone = "Not recorded"
        main_problem = "Not recorded"
    else:
        project_name = cell_text(summary_row["Project"]) or fallback_name
        client = cell_text(summary_row["Client"]) or "Not recorded"
        status = cell_text(summary_row["Status"]) or "Not recorded"
        milestone = format_milestone(
            cell_text(summary_row["Next Milestone"]),
            cell_text(summary_row["Date"]),
        )
        main_problem = cell_text(summary_row["Main Problem"]) or "Not recorded"

    label_width = len("Next Milestone")
    lines = [
        field_line("Project Name", project_name, label_width).lstrip(),
        field_line("Client", client, label_width).lstrip(),
        field_line("Status", status, label_width).lstrip(),
        field_line("Next Milestone", milestone, label_width).lstrip(),
        field_line("Main Problem", main_problem, label_width).lstrip(),
        "",
    ]

    # Split the log by Type. Comparison is case-insensitive.
    type_key = log_rows["Type"].map(lambda value: cell_text(value).casefold()) if not log_rows.empty else pd.Series(dtype=str)
    issues = log_rows.loc[type_key == ISSUE_TYPE] if not log_rows.empty else log_rows
    risks = log_rows.loc[type_key == RISK_TYPE] if not log_rows.empty else log_rows
    other = log_rows.loc[~type_key.isin([ISSUE_TYPE, RISK_TYPE])] if not log_rows.empty else log_rows

    lines.extend(section_block("Issues", issues, issue_lines))
    lines.append("")
    lines.extend(section_block("Risks", risks, risk_lines))
    lines.append("")

    # Keep change requests and any other types visible instead of dropping them.
    if not other.empty:
        for type_name, group in other.groupby(other["Type"].map(cell_text), sort=False):
            heading = type_name or "Other"
            lines.extend(section_block(heading, group, other_lines))
            lines.append("")

    lines.append("Recommended Actions")
    actions = [
        line
        for line in (recommended_action_line(row) for _, row in log_rows.iterrows())
        if line
    ]
    if actions:
        lines.extend(actions)
    else:
        lines.append("  None")

    return lines


def ordered_projects(summary: pd.DataFrame, log: pd.DataFrame) -> list[tuple[str, str, pd.Series | None]]:
    """List projects in Summary order, then any log-only projects not on Summary."""
    projects: list[tuple[str, str, pd.Series | None]] = []
    seen: set[str] = set()

    for _, row in summary.iterrows():
        name = cell_text(row["Project"])
        key = project_key(name)
        if not key or key in seen:
            continue
        seen.add(key)
        projects.append((key, name, row))

    for _, row in log.iterrows():
        name = cell_text(row["Project"])
        key = project_key(name)
        if not key or key in seen:
            continue
        seen.add(key)
        projects.append((key, name, None))

    return projects


def build_report(summary: pd.DataFrame, log: pd.DataFrame, source_name: str) -> str:
    """Compose the full weekly report as a single string."""
    generated = datetime.now().strftime("%A, %d %B %Y")
    projects = ordered_projects(summary, log)

    issue_count = 0
    risk_count = 0
    if not log.empty:
        types = log["Type"].map(lambda value: cell_text(value).casefold())
        issue_count = int((types == ISSUE_TYPE).sum())
        risk_count = int((types == RISK_TYPE).sum())

    header = [
        "TPM WEEKLY REPORT",
        "=" * 72,
        f"Generated : {generated}",
        f"Source    : {source_name}",
        f"Projects  : {len(projects)}",
        f"Issues    : {issue_count}",
        f"Risks     : {risk_count}",
        "",
    ]

    body: list[str] = []
    if not projects:
        body.append("No projects were found in the workbook.")
    else:
        for index, (key, name, summary_row) in enumerate(projects, start=1):
            project_rows = rows_for_project(log, key)
            section = build_project_section(summary_row, project_rows, name)
            # Numbered heading uses the display name, then the field block.
            heading = f"{index}. {name}"
            body.append(heading)
            body.append("-" * len(heading))
            body.extend(section)
            body.append("")
            body.append("=" * 72)
            body.append("")

    return "\n".join(header + body).rstrip() + "\n"


def generate(workbook: Path, output: Path) -> Path:
    """Load both sheets, write Weekly_Report.txt, and return the output path."""
    summary = rename_columns(load_sheet(workbook, SUMMARY_SHEET), SUMMARY_COLUMNS, SUMMARY_SHEET)
    log = rename_columns(load_sheet(workbook, RISK_SHEET), RISK_COLUMNS, RISK_SHEET)
    report = build_report(summary, log, workbook.name)

    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(report, encoding="utf-8")
    return output


def main() -> int:
    """CLI entry point. Prints the output path, or an error, and returns a status code."""
    args = parse_args()
    try:
        written = generate(args.input.resolve(), args.output.resolve())
    except (FileNotFoundError, ValueError, OSError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Wrote {written}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
