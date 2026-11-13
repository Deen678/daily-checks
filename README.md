# Daily Checks

Daily checks on the websites we look after.

## Site status checker
Reads the sites spreadsheet (Excel or CSV), works out which column holds the
website and which holds its status, checks every **Live** and **WIP** site,
and reports which are **Up**, **Down** or have a **Problem**. Sites that are
**In design** or **In dev** aren't live yet, so they're skipped.

### Use
Open this repo in Claude Code, attach the sheet and ask:

> Use the site-status agent to check this sheet

You get back a copy of the sheet with colour-coded results added (Result,
Details, HTTP code, Final address, Load time, Checked at), plus a Summary tab
listing what needs attention: Down first, then Problem. The original file is
never changed. Results are saved in `status/`.

### Rules
`reference/status-rules.json` lists which statuses are checked and which are
skipped. A status in neither list is reported as "Unknown status" so nothing
is skipped silently. Ask Claude to change the rules any time, e.g. "also check
Staging" or "stop checking WIP".

A site is only reported **Down** after failing twice. Bot protection (e.g.
Cloudflare) is reported as "Couldn't check", never as down.

## Pieces
- `.claude/agents/site-status.md`: the agent's instructions
- `scripts/status_check.py`: reads the sheet, checks the sites, writes the results workbook
- `reference/status-rules.json`: which statuses get checked

Run it directly: `python3 scripts/status_check.py sites.xlsx` (add `--only-columns` to just see which columns it would use).
Requires Python 3 with `requests` and `openpyxl`.
