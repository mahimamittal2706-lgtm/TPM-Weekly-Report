"""Build a TPM weekly status report from an Excel workbook.

Reads the Summary sheet and the Risk and Issue sheet, then writes
Weekly_Report.txt. The report opens with an executive summary, then one
section per project: client, status, next milestone, issues, risks, and
recommended actions.
"""

from __future__ import annotations

import argparse
import sys
import textwrap
from dataclasses import dataclass
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

# Weekday names used to order due dates that are stored as day names.
WEEKDAYS = (
    "monday",
    "tuesday",
    "wednesday",
    "thursday",
    "friday",
    "saturday",
    "sunday",
)
# Lower rank is more urgent. Unknown severities sort after these.
SEVERITY_RANK = {"critical": 0, "high": 1, "medium": 2, "low": 3}
# Portfolio narrative lists these statuses first, then anything else.
STATUS_ORDER = ("Red", "Amber", "Yellow", "Green")
LINE_WIDTH = 72


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
    """Format one issue with id, description, owner, action, and due date."""
    width = len("Description")
    return [
        field_line("Issue ID", cell_text(row["ID"]) or "—", width),
        field_line("Description", cell_text(row["Description"]) or "No description", width),
        field_line("Owner", cell_text(row["Owner"]) or "—", width),
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


def due_rank(due: str) -> tuple:
    """Sort Monday before Tuesday, real dates after weekday names, blanks last."""
    text = due.casefold().strip()
    for index, day in enumerate(WEEKDAYS):
        if day in text.split():
            return (0, index, text)
    parsed = pd.to_datetime(due, errors="coerce", dayfirst=False)
    if pd.notna(parsed):
        return (1, int(parsed.toordinal()), text)
    if not text or text in {"—", "no due date", "not recorded"}:
        return (3, 0, text)
    return (2, 0, text)


def severity_rank(value: str) -> int:
    """Map High / Medium / Low onto a sort rank. Blank values sort last."""
    return SEVERITY_RANK.get(value.casefold(), len(SEVERITY_RANK))


def action_sort_key(row: pd.Series, index: int, project_order: dict[str, int] | None = None) -> tuple:
    """Order actions by due date, then impact, then portfolio order.

    Blank probability does not demote a row. Issues have no probability, so a
    blank is treated as equal to the impact and a high-impact issue stays with
    high-impact risks.
    """
    impact = cell_text(row["Impact"])
    probability = cell_text(row["Probability"])
    probability_rank = severity_rank(probability) if probability else severity_rank(impact)
    if project_order is None:
        project_rank = 0
    else:
        project_rank = project_order.get(cell_text(row["Project"]).casefold(), len(project_order))
    # Project follows the due date so one day's work stays grouped by project.
    # Impact still orders items inside that project.
    return (
        due_rank(cell_text(row["Due Date"])),
        project_rank,
        severity_rank(impact),
        probability_rank,
        index,
    )


def sorted_log(log_rows: pd.DataFrame, project_order: dict[str, int] | None = None) -> pd.DataFrame:
    """Return log rows ordered for a week plan. Empty frames pass through."""
    if log_rows.empty:
        return log_rows
    ranked = sorted(
        enumerate(log_rows.to_dict(orient="records")),
        key=lambda pair: action_sort_key(pd.Series(pair[1]), pair[0], project_order),
    )
    return pd.DataFrame([record for _, record in ranked], columns=log_rows.columns)


def wrap_lines(text: str, indent: str = "  ", width: int = LINE_WIDTH) -> list[str]:
    """Wrap a prose sentence so the executive summary stays inside the rule width."""
    return textwrap.wrap(
        text,
        width=width,
        initial_indent=indent,
        subsequent_indent=indent,
    )


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
        field_line("Client Name", client, label_width).lstrip(),
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
    # Week order: Monday actions before Tuesday, high impact before low.
    actions = [
        line
        for line in (recommended_action_line(row) for _, row in sorted_log(log_rows).iterrows())
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


@dataclass
class ProjectBrief:
    """Fields the executive summary needs from one project, already display-ready."""

    name: str
    client: str
    status: str
    milestone: str
    main_problem: str
    issue_count: int
    risk_count: int


def count_phrase(count: int, singular: str) -> str:
    """'1 issue' or '4 issues'."""
    if count == 1:
        return f"1 {singular}"
    return f"{count} {singular}s"


def join_names(names: list[str]) -> str:
    """Join a short list with commas and a final 'and'."""
    if not names:
        return ""
    if len(names) == 1:
        return names[0]
    if len(names) == 2:
        return f"{names[0]} and {names[1]}"
    return ", ".join(names[:-1]) + f", and {names[-1]}"


def log_type(row: pd.Series) -> str:
    """Case-folded Type cell, used to tell issues, risks, and other rows apart."""
    return cell_text(row["Type"]).casefold()


def split_counts(log_rows: pd.DataFrame) -> tuple[int, int]:
    """Return issue and risk counts for one project, or for the whole log."""
    if log_rows.empty:
        return 0, 0
    types = log_rows["Type"].map(lambda value: cell_text(value).casefold())
    return int((types == ISSUE_TYPE).sum()), int((types == RISK_TYPE).sum())


def make_brief(name: str, summary_row: pd.Series | None, log_rows: pd.DataFrame) -> ProjectBrief:
    """Collect the summary fields and open-item counts for one project."""
    if summary_row is None:
        client = "Not recorded"
        status = "Not recorded"
        milestone = "Not recorded"
        main_problem = "Not recorded"
    else:
        client = cell_text(summary_row["Client"]) or "Not recorded"
        status = cell_text(summary_row["Status"]) or "Not recorded"
        milestone = format_milestone(
            cell_text(summary_row["Next Milestone"]),
            cell_text(summary_row["Date"]),
        )
        main_problem = cell_text(summary_row["Main Problem"]) or "Not recorded"
    issue_count, risk_count = split_counts(log_rows)
    return ProjectBrief(
        name=name,
        client=client,
        status=status,
        milestone=milestone,
        main_problem=main_problem,
        issue_count=issue_count,
        risk_count=risk_count,
    )


def exposure_phrase(row: pd.Series) -> str:
    """Short probability and impact note for the week-at-a-glance list."""
    probability = cell_text(row["Probability"])
    impact = cell_text(row["Impact"])
    parts = []
    if probability:
        parts.append(f"{probability} probability")
    if impact:
        parts.append(f"{impact} impact")
    return ", ".join(parts)


def portfolio_headline(briefs: list[ProjectBrief], log: pd.DataFrame, issue_count: int, risk_count: int) -> str:
    """One opening sentence: how many projects, how they sit, how much is open."""
    if not briefs:
        return "No projects were found in the workbook."

    named: list[str] = []
    seen: set[str] = set()
    for status in STATUS_ORDER:
        matches = [brief.name for brief in briefs if brief.status.casefold() == status.casefold()]
        if matches:
            seen.add(status.casefold())
            named.append(f"{len(matches)} {status} ({join_names(matches)})")
    for brief in briefs:
        key = brief.status.casefold()
        if key in seen:
            continue
        matches = [item.name for item in briefs if item.status.casefold() == key]
        seen.add(key)
        named.append(f"{len(matches)} {brief.status} ({join_names(matches)})")

    extra = ""
    if not log.empty:
        types = log["Type"].map(lambda value: cell_text(value).casefold())
        other = log.loc[~types.isin([ISSUE_TYPE, RISK_TYPE])]
        if not other.empty:
            labels = []
            for label in other["Type"].map(cell_text):
                if label and label not in labels:
                    labels.append(label)
            if len(labels) == 1:
                extra = f", plus {count_phrase(len(other), labels[0].lower())}"
            else:
                extra = f", plus {count_phrase(len(other), 'other item')}"

    return (
        f"{count_phrase(len(briefs), 'project').capitalize()} this week: {join_names(named)}. "
        f"{count_phrase(issue_count, 'issue').capitalize()} and {count_phrase(risk_count, 'risk')} "
        f"are open{extra}."
    )


def project_paragraph(brief: ProjectBrief) -> str:
    """Status, milestone, main problem, and open-item counts for one project."""
    client = brief.client if brief.client != "Not recorded" else "client not recorded"
    sentence = f"{brief.name} ({client}) is {brief.status}. Next milestone: {brief.milestone}."
    if brief.main_problem and brief.main_problem != "Not recorded":
        problem = brief.main_problem
        if problem[0].isalpha():
            problem = problem[0].lower() + problem[1:]
        sentence += f" Main problem: {problem}."
    sentence += (
        f" Open this week: {count_phrase(brief.issue_count, 'issue')} and "
        f"{count_phrase(brief.risk_count, 'risk')}."
    )
    return sentence


def week_at_a_glance(log: pd.DataFrame, project_order: dict[str, int] | None = None) -> list[str]:
    """Actions grouped by due date, Monday first, high impact before low impact."""
    lines = ["Week at a glance"]
    if log.empty:
        lines.append("  None")
        return lines

    groups: list[tuple[str, list[pd.Series]]] = []
    index_by_due: dict[str, int] = {}
    for _, row in sorted_log(log, project_order).iterrows():
        if not cell_text(row["Action"]):
            continue
        label = cell_text(row["Due Date"]) or "No due date"
        key = label.casefold()
        if key not in index_by_due:
            index_by_due[key] = len(groups)
            groups.append((label, []))
        groups[index_by_due[key]][1].append(row)

    if not groups:
        lines.append("  None")
        return lines

    for label, rows in groups:
        lines.append(f"  {label}")
        for row in rows:
            item_id = cell_text(row["ID"]) or "—"
            action = cell_text(row["Action"])
            owner = cell_text(row["Owner"]) or "Unassigned"
            project = cell_text(row["Project"]) or "—"
            kind = cell_text(row["Type"]) or "Item"
            detail = f"    - [{item_id}] {action} — {owner} — {project} {kind}"
            exposure = exposure_phrase(row)
            if exposure:
                detail += f" — {exposure}"
            lines.append(detail)
    return lines


def owner_conflicts(log: pd.DataFrame, project_order: dict[str, int] | None = None) -> list[str]:
    """Owners who have actions on two projects with the same due date."""
    if log.empty:
        return []

    buckets: dict[tuple[str, str], list[pd.Series]] = {}
    order: list[tuple[str, str]] = []
    for _, row in sorted_log(log, project_order).iterrows():
        owner = cell_text(row["Owner"])
        if not cell_text(row["Action"]) or not owner:
            continue
        due = cell_text(row["Due Date"]) or "No due date"
        key = (owner.casefold(), due.casefold())
        if key not in buckets:
            buckets[key] = []
            order.append(key)
        buckets[key].append(row)

    conflicts = []
    for key in order:
        rows = buckets[key]
        projects: list[str] = []
        for row in rows:
            project = cell_text(row["Project"])
            if project and project not in projects:
                projects.append(project)
        if len(projects) < 2:
            continue
        owner = cell_text(rows[0]["Owner"])
        due = cell_text(rows[0]["Due Date"]) or "No due date"
        ids = ", ".join(cell_text(row["ID"]) or "—" for row in rows)
        conflicts.append(
            f"{owner} is on {due} for both {join_names(projects)} ({ids}). Sequence that work before both slip."
        )
    return conflicts


def elevated_risk_bullets(log: pd.DataFrame) -> list[str]:
    """Risks with high probability and medium or high impact, highest exposure first."""
    if log.empty:
        return []
    bullets = []
    for _, row in sorted_log(log).iterrows():
        if log_type(row) != RISK_TYPE:
            continue
        probability = cell_text(row["Probability"])
        impact = cell_text(row["Impact"])
        if severity_rank(probability) > SEVERITY_RANK["high"]:
            continue
        if severity_rank(impact) > SEVERITY_RANK["medium"]:
            continue
        item_id = cell_text(row["ID"]) or "—"
        project = cell_text(row["Project"]) or "the portfolio"
        description = cell_text(row["Description"]) or "an open risk"
        owner = cell_text(row["Owner"]) or "Unassigned"
        due = cell_text(row["Due Date"]) or "no due date"
        if severity_rank(impact) <= SEVERITY_RANK["high"]:
            exposure = "Probability and impact are both High."
        else:
            exposure = f"Probability is High and impact is {impact}."
        bullets.append(
            f"Mitigate {project} {item_id} before {due}: {description} ({owner}). {exposure}"
        )
    return bullets


def earliest_due_bullet(log: pd.DataFrame, project_order: dict[str, int] | None = None) -> str | None:
    """Point at the first due date so Monday work is not buried under later items."""
    if log.empty:
        return None
    first_due = ""
    ids: list[str] = []
    for _, row in sorted_log(log, project_order).iterrows():
        if not cell_text(row["Action"]):
            continue
        due = cell_text(row["Due Date"]) or "No due date"
        if not first_due:
            first_due = due
        if due.casefold() != first_due.casefold():
            break
        ids.append(cell_text(row["ID"]) or "—")
    if not first_due or not ids:
        return None
    return f"Clear {first_due} first ({', '.join(ids)}). Later due dates depend on that work landing."


def focus_bullets(log: pd.DataFrame, project_order: dict[str, int] | None = None) -> list[str]:
    """The few actions leadership should track, capped so the summary stays short."""
    bullets = elevated_risk_bullets(log)
    bullets.extend(owner_conflicts(log, project_order))
    earliest = earliest_due_bullet(log, project_order)
    if earliest:
        bullets.append(earliest)
    if not bullets:
        return ["No high-exposure risks or cross-project clashes are open. Work actions in due-date order."]
    return bullets[:6]


def wrap_bullet(number: int, text: str) -> list[str]:
    """Wrap a numbered focus item so continuation lines line up under the text."""
    prefix = f"  {number}. "
    return textwrap.wrap(
        text,
        width=LINE_WIDTH,
        initial_indent=prefix,
        subsequent_indent=" " * len(prefix),
    )


def build_executive_summary(
    briefs: list[ProjectBrief],
    log: pd.DataFrame,
    issue_count: int,
    risk_count: int,
) -> list[str]:
    """Portfolio headline, one paragraph per project, the week plan, and focus items."""
    lines = ["EXECUTIVE SUMMARY", "-" * LINE_WIDTH]
    lines.extend(wrap_lines(portfolio_headline(briefs, log, issue_count, risk_count), indent=""))
    lines.append("")

    if briefs:
        lines.append("Portfolio")
        for brief in briefs:
            lines.extend(wrap_lines(project_paragraph(brief)))
        lines.append("")

    project_order = {brief.name.casefold(): index for index, brief in enumerate(briefs)}
    lines.extend(week_at_a_glance(log, project_order))
    lines.append("")
    lines.append("Recommended focus")
    for number, bullet in enumerate(focus_bullets(log, project_order), start=1):
        lines.extend(wrap_bullet(number, bullet))
    return lines


def build_report(summary: pd.DataFrame, log: pd.DataFrame, source_name: str) -> str:
    """Compose the full weekly report as a single string."""
    generated = datetime.now().strftime("%A, %d %B %Y")
    projects = ordered_projects(summary, log)

    issue_count, risk_count = split_counts(log)

    header = [
        "TPM WEEKLY REPORT",
        "=" * LINE_WIDTH,
        f"Generated : {generated}",
        f"Source    : {source_name}",
        f"Projects  : {len(projects)}",
        f"Issues    : {issue_count}",
        f"Risks     : {risk_count}",
        "",
    ]

    briefs: list[ProjectBrief] = []
    body: list[str] = []
    if not projects:
        body.append("No projects were found in the workbook.")
    else:
        for index, (key, name, summary_row) in enumerate(projects, start=1):
            project_rows = rows_for_project(log, key)
            briefs.append(make_brief(name, summary_row, project_rows))
            section = build_project_section(summary_row, project_rows, name)
            # Numbered heading uses the display name, then the field block.
            heading = f"{index}. {name}"
            body.append(heading)
            body.append("-" * len(heading))
            body.extend(section)
            body.append("")
            body.append("=" * LINE_WIDTH)
            body.append("")

    summary_lines = build_executive_summary(briefs, log, issue_count, risk_count)
    parts = header + summary_lines + ["", "=" * LINE_WIDTH, ""] + body
    return "\n".join(parts).rstrip() + "\n"


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
