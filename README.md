# TPM Weekly Report

Python tool that reads `TPM.xlsx` and writes a plain-text weekly status report.

It uses two sheets:

- **Summary** — project, client, status, next milestone, milestone date, and main problem
- **Risk and Issue** — issues, risks, change requests, owners, actions, and due dates

The report opens with an executive summary: each project's status and next milestone, the highest-exposure risk, the week's actions by due date, and a short recommended-focus list.

Each project in `Weekly_Report.txt` then includes:

- Project Name
- Client Name
- Status
- Next Milestone
- Issues (ID, description, owner, action, due date)
- Risks (ID, description, probability, impact, owner, mitigation action, due date)
- Recommended Actions

Change requests (and any other type that is not Issue or Risk) are listed in their own section so they are not dropped. Recommended actions are taken from the Action column and ordered by due date.

## Setup

Requires Python 3.10 or newer.

```powershell
python -m pip install -r requirements.txt
```

## Run

From this folder, with `TPM.xlsx` in the same directory:

```powershell
python report_generator.py
```

That writes `Weekly_Report.txt` next to the script.

Custom paths:

```powershell
python report_generator.py --input TPM.xlsx --output Weekly_Report.txt
```

## Workbook columns

**Summary:** Project, Client, Status, Next Milestone, Date, Main Problem

**Risk and Issue:** ID, Project, Type, Description, Probability, Impact, Owner, Action, Due Date

Project names are matched between sheets without regard to case or extra spaces. Projects appear in Summary order. A project that exists only on Risk and Issue is still included, with client and status marked as not recorded.
