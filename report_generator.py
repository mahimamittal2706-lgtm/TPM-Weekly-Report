"""Build a TPM weekly status report from an Excel workbook.

Reads the Summary sheet and the Risk and Issue sheet, then writes
Weekly_Report.txt. Every project, issue, risk, and other logged row is
included. Fields the tracker does not contain are shown as "Not recorded".
"""

from __future__ import annotations

import argparse
import sys
import textwrap
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

# Printed when the user asked for a field the workbook does not have.
NOT_RECORDED = "Not recorded"

ISSUE_TYPE = "issue"
RISK_TYPE = "risk"

WEEKDAYS = (
    "monday",
    "tuesday",
    "wednesday",
    "thursday",
    "friday",
    "saturday",
    "sunday",
)
SEVERITY_RANK = {"critical": 0, "high": 1, "medium": 2, "low": 3}
# Worst status first. Amber and Yellow are the same tier; Amber is listed first.
HEALTH_ORDER = ("red", "amber", "yellow", "green")
LINE_WIDTH = 78


def parse_args() -> argparse.Namespace:
    """Accept optional workbook and output paths. Defaults sit beside this script."""
    script_dir = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(
        description="Generate a TPM weekly status report from Summary and Risk and Issue sheets."
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
        raise ValueError(
            f"Sheet '{sheet_name}' was not found in {workbook.name}."
        ) from exc

    frame = frame.dropna(how="all")
    frame.columns = [str(column).strip() for column in frame.columns]
    return frame


def rename_columns(frame: pd.DataFrame, mapping: dict[str, str], sheet_name: str) -> pd.DataFrame:
    """Rename required headers and keep every other column so no tracker field is dropped."""
    lookup = {str(column).lower(): column for column in frame.columns}
    missing = [canonical for key, canonical in mapping.items() if key not in lookup]
    if missing:
        raise ValueError(
            f"Sheet '{sheet_name}' is missing columns: {', '.join(missing)}."
        )

    renamed = frame.rename(columns={lookup[key]: canonical for key, canonical in mapping.items()})
    ordered = list(mapping.values())
    extras = [column for column in renamed.columns if column not in ordered]
    return renamed.loc[:, ordered + extras].copy()


def cell_text(value: object) -> str:
    """Turn a cell into display text. Blank cells and NaN become an empty string."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    if isinstance(value, pd.Timestamp):
        return value.strftime("%d %b %Y")
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def project_key(name: object) -> str:
    """Key used to join Summary rows with Risk and Issue rows."""
    return cell_text(name).casefold()


def column_text(row: pd.Series, column: str) -> str:
    """Read one column when it exists. A missing column is an empty string."""
    if column not in row.index:
        return ""
    return cell_text(row[column])


def show(value: str) -> str:
    """Display text for a requested field. Blank tracker cells stay explicit."""
    return value if value else NOT_RECORDED


def due_rank(due: str) -> tuple:
    """Sort Monday before Tuesday, real dates after weekday names, blanks last."""
    text = due.casefold().strip()
    for index, day in enumerate(WEEKDAYS):
        if day in text.split():
            return (0, index, text)
    parsed = pd.to_datetime(due, errors="coerce", dayfirst=False)
    if pd.notna(parsed):
        return (1, int(parsed.toordinal()), text)
    if not text or text.casefold() in {"—", "no due date", "not recorded"}:
        return (3, 0, text)
    return (2, 0, text)


def severity_rank(value: str) -> int:
    """Map Critical / High / Medium / Low onto a sort rank. Blank values sort last."""
    return SEVERITY_RANK.get(value.casefold(), len(SEVERITY_RANK))


def log_type(row: pd.Series) -> str:
    """Case-folded Type cell, used to tell issues, risks, and other rows apart."""
    return column_text(row, "Type").casefold()


def split_by_type(log: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Split the log into issues, risks, and every remaining type."""
    if log.empty:
        empty = log
        return empty, empty, empty
    types = log["Type"].map(lambda value: cell_text(value).casefold())
    issues = log.loc[types == ISSUE_TYPE]
    risks = log.loc[types == RISK_TYPE]
    other = log.loc[~types.isin([ISSUE_TYPE, RISK_TYPE])]
    return issues, risks, other


def rows_for_project(log: pd.DataFrame, key: str) -> pd.DataFrame:
    """Filter the log to one project, preserving sheet order."""
    if log.empty:
        return log
    mask = log["Project"].map(project_key) == key
    return log.loc[mask]


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

    if not log.empty:
        for _, row in log.iterrows():
            name = cell_text(row["Project"])
            key = project_key(name)
            if not key or key in seen:
                continue
            seen.add(key)
            projects.append((key, name, None))

    return projects


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


def health_rank(status: str) -> int:
    """Lower rank is a weaker RAG status. Unrecognised statuses sort after Green."""
    folded = status.casefold()
    if folded in HEALTH_ORDER:
        return HEALTH_ORDER.index(folded)
    if folded in {"", "not recorded"}:
        return len(HEALTH_ORDER) + 1
    return len(HEALTH_ORDER)


def overall_health(statuses: list[str]) -> str:
    """Portfolio health is the weakest recorded RAG status."""
    recorded = [status for status in statuses if status and status.casefold() != "not recorded"]
    if not recorded:
        return NOT_RECORDED
    worst = min(recorded, key=health_rank)
    canonical = {"red": "Red", "amber": "Amber", "yellow": "Yellow", "green": "Green"}
    return canonical.get(worst.casefold(), worst)


def is_elevated(row: pd.Series) -> bool:
    """High impact, or high probability with medium-or-higher impact, from logged cells."""
    impact = severity_rank(column_text(row, "Impact"))
    probability = severity_rank(column_text(row, "Probability"))
    if impact <= SEVERITY_RANK["high"]:
        return True
    return probability <= SEVERITY_RANK["high"] and impact <= SEVERITY_RANK["medium"]


def is_escalation(action: str) -> bool:
    """True when the logged action itself says to escalate."""
    return "escalat" in action.casefold()


def is_pending_decision(action: str) -> bool:
    """True when the logged action is to assess, decide, or obtain approval."""
    text = action.casefold()
    markers = ("assess", "decide", "decision", "approve", "approval")
    return any(marker in text for marker in markers)


def classify_action(action: str) -> str:
    """Bucket a logged action. Escalation wins when the wording matches both rules."""
    if is_escalation(action):
        return "escalation"
    if is_pending_decision(action):
        return "decision"
    return "immediate"


def rule(title: str) -> list[str]:
    """A major section heading and its underline."""
    return [title, "-" * LINE_WIDTH]


def wrap_text(text: str, indent: str = "", subsequent: str | None = None) -> list[str]:
    """Wrap prose to the report width."""
    return textwrap.wrap(
        text,
        width=LINE_WIDTH,
        initial_indent=indent,
        subsequent_indent=indent if subsequent is None else subsequent,
    ) or [indent.rstrip()]


def field_block(pairs: list[tuple[str, str]]) -> list[str]:
    """Align labelled fields and wrap long values under the value column."""
    if not pairs:
        return []
    width = max(len(label) for label, _ in pairs)
    lines: list[str] = []
    for label, value in pairs:
        shown = show(value)
        prefix = f"  {label:<{width}} : "
        wrapped = textwrap.wrap(shown, width=max(24, LINE_WIDTH - len(prefix))) or [shown]
        lines.append(prefix + wrapped[0])
        hang = " " * len(prefix)
        for extra in wrapped[1:]:
            lines.append(hang + extra)
    return lines


def extra_pairs(row: pd.Series, consumed: set[str]) -> list[tuple[str, str]]:
    """Any tracker column that is not already printed, when the cell has a value."""
    pairs = []
    for column in row.index:
        if column in consumed:
            continue
        text = cell_text(row[column])
        if text:
            pairs.append((str(column), text))
    return pairs


def exposure_text(row: pd.Series) -> str:
    """Probability and impact exactly as stored."""
    parts = []
    probability = column_text(row, "Probability")
    impact = column_text(row, "Impact")
    if probability:
        parts.append(f"probability {probability}")
    if impact:
        parts.append(f"impact {impact}")
    return ", ".join(parts)


def project_fields(name: str, summary_row: pd.Series | None) -> dict[str, str]:
    """Display fields for one project. Missing summary cells stay 'Not recorded'."""
    if summary_row is None:
        return {
            "name": name,
            "client": "",
            "status": "",
            "progress": "",
            "milestone": "",
            "milestone_date": "",
            "main_problem": "",
        }
    return {
        "name": cell_text(summary_row["Project"]) or name,
        "client": cell_text(summary_row["Client"]),
        "status": cell_text(summary_row["Status"]),
        "progress": column_text(summary_row, "Progress"),
        "milestone": cell_text(summary_row["Next Milestone"]),
        "milestone_date": cell_text(summary_row["Date"]),
        "main_problem": cell_text(summary_row["Main Problem"]),
    }


def summary_extra_pairs(summary_row: pd.Series | None) -> list[tuple[str, str]]:
    """Summary columns beyond the standard set, so a new tracker column is still printed."""
    if summary_row is None:
        return []
    consumed = {
        "Project",
        "Client",
        "Status",
        "Next Milestone",
        "Date",
        "Main Problem",
        "Progress",
    }
    return extra_pairs(summary_row, consumed)


def build_executive_summary(
    projects: list[tuple[str, str, pd.Series | None]],
    log: pd.DataFrame,
) -> list[str]:
    """Counts, portfolio health, and critical concerns drawn from tracker cells."""
    issues, risks, other = split_by_type(log)
    fields = [project_fields(name, row) for _, name, row in projects]
    statuses = [item["status"] for item in fields]
    health = overall_health(statuses)

    lines = rule("EXECUTIVE SUMMARY")
    lines.extend(
        field_block(
            [
                ("Total projects", str(len(projects))),
                ("Overall project health", health),
                ("Total open issues", str(len(issues))),
                ("Total risks", str(len(risks))),
            ]
        )
    )
    lines.append("")

    if fields:
        status_bits = []
        for item in fields:
            client = f" ({item['client']})" if item["client"] else ""
            status_bits.append(f"{item['name']}{client} is {show(item['status'])}")
        health_sentence = (
            f"Portfolio health is {health}. {join_names(status_bits)}. "
            f"The open log holds {count_phrase(len(issues), 'issue')} and "
            f"{count_phrase(len(risks), 'risk')}"
        )
        if not other.empty:
            labels = []
            for label in other["Type"].map(cell_text):
                shown = label or "untyped item"
                if shown not in labels:
                    labels.append(shown)
            health_sentence += f", plus {count_phrase(len(other), labels[0] if len(labels) == 1 else 'other item')}"
        health_sentence += (
            ". Issue rows are counted as open because the tracker has no closed flag "
            "and no issue-status column. Progress and issue priority are not in the tracker."
        )
        lines.extend(wrap_text(health_sentence))
        lines.append("")

    lines.append("Critical concerns")
    concerns = critical_concerns(fields, issues, risks)
    if not concerns:
        lines.append("  None recorded in the tracker.")
    else:
        for number, text in enumerate(concerns, start=1):
            prefix = f"  {number}. "
            lines.extend(
                textwrap.wrap(
                    text,
                    width=LINE_WIDTH,
                    initial_indent=prefix,
                    subsequent_indent=" " * len(prefix),
                )
            )
    return lines


def critical_concerns(
    fields: list[dict[str, str]],
    issues: pd.DataFrame,
    risks: pd.DataFrame,
) -> list[str]:
    """Red, Amber, and Yellow projects, then elevated issues and risks."""
    concerns = []
    for item in fields:
        if health_rank(item["status"]) > HEALTH_ORDER.index("yellow"):
            continue
        milestone = item["milestone"] or NOT_RECORDED
        due = item["milestone_date"] or NOT_RECORDED
        client = item["client"] or NOT_RECORDED
        problem = item["main_problem"] or NOT_RECORDED
        concerns.append(
            f"{item['name']} ({client}) is {item['status']}. "
            f"Main problem: {problem}. Next milestone: {milestone}, due {due}."
        )

    elevated_rows = []
    for frame in (issues, risks):
        if frame.empty:
            continue
        for index, (_, row) in enumerate(frame.iterrows()):
            if is_elevated(row):
                elevated_rows.append((row, index))

    elevated_rows.sort(
        key=lambda pair: (
            severity_rank(column_text(pair[0], "Impact")),
            severity_rank(column_text(pair[0], "Probability")),
            due_rank(column_text(pair[0], "Due Date")),
            pair[1],
        )
    )
    for row, _ in elevated_rows:
        item_id = column_text(row, "ID") or "—"
        project = column_text(row, "Project") or "Unassigned project"
        kind = column_text(row, "Type") or "Item"
        description = column_text(row, "Description") or "No description"
        owner = column_text(row, "Owner") or "Unassigned"
        action = column_text(row, "Action") or "no action recorded"
        due = column_text(row, "Due Date") or "no due date"
        exposure = exposure_text(row) or "severity not recorded"
        exposure = exposure[0].upper() + exposure[1:]
        concerns.append(
            f"{project} {item_id} ({kind}): {description}. {exposure}. "
            f"Owner {owner}. Action required: {action}. Due {due}."
        )
    return concerns


def build_project_status(projects: list[tuple[str, str, pd.Series | None]], log: pd.DataFrame) -> list[str]:
    """One block per project, including projects that appear only on the log."""
    lines = rule("PROJECT-WISE STATUS")
    if not projects:
        lines.append("No projects were found in the workbook.")
        return lines

    for index, (key, name, summary_row) in enumerate(projects, start=1):
        if index > 1:
            lines.append("")
        item = project_fields(name, summary_row)
        project_rows = rows_for_project(log, key)
        issues, risks, other = split_by_type(project_rows)
        heading = f"{index}. {item['name']}"
        lines.append(heading)
        lines.append("-" * len(heading))
        pairs = [
            ("Project name", item["name"]),
            ("Client name", item["client"]),
            ("Current status", item["status"]),
            ("Progress", item["progress"]),
            ("Next milestone", item["milestone"]),
            ("Milestone due date", item["milestone_date"]),
            ("Main problem", item["main_problem"]),
            ("Open issues", str(len(issues))),
            ("Open risks", str(len(risks))),
            ("Other logged items", str(len(other))),
        ]
        pairs.extend(summary_extra_pairs(summary_row))
        lines.extend(field_block(pairs))
    return lines


def issue_pairs(row: pd.Series) -> list[tuple[str, str]]:
    """Every issue field requested, plus severity columns that exist on the log."""
    pairs = [
        ("Issue ID", column_text(row, "ID")),
        ("Project", column_text(row, "Project")),
        ("Description", column_text(row, "Description")),
        ("Owner", column_text(row, "Owner")),
        ("Priority", column_text(row, "Priority")),
        ("Probability", column_text(row, "Probability")),
        ("Impact", column_text(row, "Impact")),
        ("Current status", column_text(row, "Item Status") or column_text(row, "Status")),
        ("Action required", column_text(row, "Action")),
        ("Due date", column_text(row, "Due Date")),
    ]
    pairs.extend(
        extra_pairs(
            row,
            {
                "ID",
                "Project",
                "Type",
                "Description",
                "Owner",
                "Priority",
                "Probability",
                "Impact",
                "Item Status",
                "Status",
                "Action",
                "Due Date",
            },
        )
    )
    return pairs


def risk_pairs(row: pd.Series) -> list[tuple[str, str]]:
    """Every risk field requested, plus the project so the row stays traceable."""
    pairs = [
        ("Risk ID", column_text(row, "ID")),
        ("Project", column_text(row, "Project")),
        ("Description", column_text(row, "Description")),
        ("Probability", column_text(row, "Probability")),
        ("Impact", column_text(row, "Impact")),
        ("Owner", column_text(row, "Owner")),
        ("Mitigation plan", column_text(row, "Action")),
        ("Due date", column_text(row, "Due Date")),
    ]
    pairs.extend(
        extra_pairs(
            row,
            {
                "ID",
                "Project",
                "Type",
                "Description",
                "Probability",
                "Impact",
                "Owner",
                "Action",
                "Due Date",
            },
        )
    )
    return pairs


def other_pairs(row: pd.Series) -> list[tuple[str, str]]:
    """A log row that is neither an issue nor a risk, with every stored column."""
    pairs = [
        ("ID", column_text(row, "ID")),
        ("Project", column_text(row, "Project")),
        ("Type", column_text(row, "Type")),
        ("Description", column_text(row, "Description")),
        ("Probability", column_text(row, "Probability")),
        ("Impact", column_text(row, "Impact")),
        ("Owner", column_text(row, "Owner")),
        ("Action required", column_text(row, "Action")),
        ("Due date", column_text(row, "Due Date")),
    ]
    pairs.extend(
        extra_pairs(
            row,
            {
                "ID",
                "Project",
                "Type",
                "Description",
                "Probability",
                "Impact",
                "Owner",
                "Action",
                "Due Date",
            },
        )
    )
    return pairs


def render_register(title: str, rows: pd.DataFrame, formatter, empty_line: str) -> list[str]:
    """A full register in sheet order. Empty registers say so instead of being omitted."""
    lines = rule(title)
    if rows.empty:
        lines.append(empty_line)
        return lines
    lines.append(f"Items listed: {len(rows)}")
    lines.append("")
    for index, (_, row) in enumerate(rows.iterrows(), start=1):
        if index > 1:
            lines.append("")
        lines.append(f"{index}.")
        lines.extend(field_block(formatter(row)))
    return lines


def action_entry(row: pd.Series) -> list[str]:
    """One recommended action with the tracker fields needed to act on it."""
    kind = column_text(row, "Type") or "Item"
    item_id = column_text(row, "ID") or "—"
    project = column_text(row, "Project") or "Unassigned project"
    action = column_text(row, "Action") or "No action recorded"
    pairs = [
        ("Item", f"[{item_id}] {project} {kind}"),
        ("Action", action),
        ("Description", column_text(row, "Description")),
        ("Owner", column_text(row, "Owner")),
        ("Due date", column_text(row, "Due Date")),
        ("Probability", column_text(row, "Probability")),
        ("Impact", column_text(row, "Impact")),
    ]
    return field_block(pairs)


def sort_actions(rows: list[pd.Series]) -> list[pd.Series]:
    """Due date, then impact, then probability. Original order breaks remaining ties."""
    ranked = sorted(
        enumerate(rows),
        key=lambda pair: (
            due_rank(column_text(pair[1], "Due Date")),
            severity_rank(column_text(pair[1], "Impact")),
            severity_rank(column_text(pair[1], "Probability")),
            pair[0],
        ),
    )
    return [row for _, row in ranked]


def render_action_group(title: str, rows: list[pd.Series], empty_line: str) -> list[str]:
    """One recommended-action subsection."""
    lines = [title]
    if not rows:
        lines.append(f"  {empty_line}")
        return lines
    for index, row in enumerate(sort_actions(rows), start=1):
        if index > 1:
            lines.append("")
        lines.append(f"  {index}.")
        lines.extend(action_entry(row))
    return lines


def owner_overlaps(log: pd.DataFrame) -> list[str]:
    """Owners with logged actions on two projects that share a due date."""
    if log.empty:
        return []
    buckets: dict[tuple[str, str], list[pd.Series]] = {}
    order: list[tuple[str, str]] = []
    for _, row in log.iterrows():
        owner = column_text(row, "Owner")
        if not column_text(row, "Action") or not owner:
            continue
        due = column_text(row, "Due Date") or "No due date"
        key = (owner.casefold(), due.casefold())
        if key not in buckets:
            buckets[key] = []
            order.append(key)
        buckets[key].append(row)

    overlaps = []
    for key in order:
        rows = buckets[key]
        projects: list[str] = []
        for row in rows:
            project = column_text(row, "Project")
            if project and project not in projects:
                projects.append(project)
        if len(projects) < 2:
            continue
        owner = column_text(rows[0], "Owner")
        due = column_text(rows[0], "Due Date") or "No due date"
        ids = ", ".join(column_text(row, "ID") or "—" for row in rows)
        overlaps.append(
            f"{owner} has actions on {due} for {join_names(projects)} ({ids})."
        )
    return overlaps


def build_recommended_actions(log: pd.DataFrame) -> list[str]:
    """Split logged actions into immediate work, decisions, and escalations."""
    lines = rule("RECOMMENDED ACTIONS")
    lines.extend(
        wrap_text(
            "Each group below is taken from the Action column. "
            "An action is an escalation when it says escalate. "
            "An action is a pending decision when it says assess, decide, or approve "
            "and it is not already an escalation. "
            "Every other logged action is listed under Immediate actions, in due-date order. "
            "No action row is omitted."
        )
    )
    lines.append("")

    immediate: list[pd.Series] = []
    decisions: list[pd.Series] = []
    escalations: list[pd.Series] = []
    if not log.empty:
        for _, row in log.iterrows():
            action = column_text(row, "Action")
            if not action:
                immediate.append(row)
                continue
            bucket = classify_action(action)
            if bucket == "escalation":
                escalations.append(row)
            elif bucket == "decision":
                decisions.append(row)
            else:
                immediate.append(row)

    lines.extend(
        render_action_group(
            "Immediate actions",
            immediate,
            "None recorded in the tracker.",
        )
    )
    lines.append("")
    lines.extend(
        render_action_group(
            "Pending decisions",
            decisions,
            "None recorded in the tracker.",
        )
    )
    lines.append("")
    lines.extend(
        render_action_group(
            "Escalations required",
            escalations,
            "None recorded in the tracker.",
        )
    )

    overlaps = owner_overlaps(log)
    if overlaps:
        lines.append("")
        lines.append("Same-day owner overlap")
        for text in overlaps:
            lines.extend(wrap_text(text, indent="  - ", subsequent="    "))
    return lines


def build_milestones(projects: list[tuple[str, str, pd.Series | None]]) -> list[str]:
    """Upcoming milestones in due-date order, one block per project."""
    lines = rule("UPCOMING MILESTONES")
    if not projects:
        lines.append("No projects were found in the workbook.")
        return lines

    ranked = sorted(
        enumerate(project_fields(name, row) for _, name, row in projects),
        key=lambda pair: (due_rank(pair[1]["milestone_date"]), pair[0]),
    )
    lines.append(f"Milestones listed: {len(ranked)}")
    lines.append("")
    for position, (_, item) in enumerate(ranked, start=1):
        if position > 1:
            lines.append("")
        lines.append(f"{position}. {item['name']}")
        lines.extend(
            field_block(
                [
                    ("Project name", item["name"]),
                    ("Client name", item["client"]),
                    ("Current status", item["status"]),
                    ("Next milestone", item["milestone"]),
                    ("Milestone due date", item["milestone_date"]),
                    ("Main problem", item["main_problem"]),
                ]
            )
        )
    return lines


def build_report(summary: pd.DataFrame, log: pd.DataFrame, source_name: str) -> str:
    """Compose the full weekly status report as a single string."""
    generated = datetime.now().strftime("%A, %d %B %Y")
    projects = ordered_projects(summary, log)
    issues, risks, other = split_by_type(log)
    health = overall_health(
        [project_fields(name, row)["status"] for _, name, row in projects]
    )

    header = [
        "TPM WEEKLY STATUS REPORT",
        "=" * LINE_WIDTH,
        f"Generated : {generated}",
        f"Source    : {source_name}",
        f"Sheets    : {SUMMARY_SHEET}; {RISK_SHEET}",
        f"Projects  : {len(projects)}",
        f"Health    : {health}",
        f"Issues    : {len(issues)}",
        f"Risks     : {len(risks)}",
        f"Other     : {len(other)}",
        "",
        "Fields marked \"Not recorded\" are absent from the tracker. "
        "Every project, issue, risk, and other logged row is included.",
        "",
    ]
    # Wrap the note to the rule width.
    note = header[-2]
    header = header[:-2] + wrap_text(note) + [""]

    sections = [
        build_executive_summary(projects, log),
        build_project_status(projects, log),
        render_register(
            "ISSUES REGISTER",
            issues,
            issue_pairs,
            "No issues were found in the tracker.",
        ),
        render_register(
            "RISK REGISTER",
            risks,
            risk_pairs,
            "No risks were found in the tracker.",
        ),
        render_register(
            "OTHER LOGGED ITEMS",
            other,
            other_pairs,
            "No other item types were found in the tracker.",
        ),
        build_recommended_actions(log),
        build_milestones(projects),
    ]

    parts: list[str] = header
    for section in sections:
        parts.extend(section)
        parts.append("")
        parts.append("=" * LINE_WIDTH)
        parts.append("")

    parts.append("END OF REPORT")
    parts.append(
        f"Projects {len(projects)} | Issues {len(issues)} | Risks {len(risks)} | Other {len(other)}"
    )
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
