# Premiership Rugby Analytics Workbook

A standalone Microsoft Excel tool for tracking English Premiership Rugby
team and player statistics, built to hold a historic archive back to the
2000/01 season alongside the upcoming 2026/27 season.

## Files

- `Premiership_Rugby_Analytics.xlsx` — the workbook. Open it and start with
  the **README** sheet inside.
- `build_workbook.py` — the script that generates the workbook (uses
  `openpyxl`). Kept for reproducibility; you don't need it to use the file.

## How it works

The workbook does not auto-scrape the web — modern rugby stats sites render
their tables with JavaScript, which Excel's built-in web tools can't read.
Instead it uses a simple copy/paste-and-clean workflow:

1. Visit a recommended source page (see the **Sources** sheet).
2. Copy the table you see on screen and paste it into the matching
   `...RawPaste` sheet.
3. Copy the "ready to archive" block next to the paste zone and paste it
   into the matching `...Archive` table, which keeps permanent history.

A **Season_Explorer** sheet lets you pick any season and see its full
standings, and a **Dashboard** sheet charts a team's league points across
every season plus the top 10 try scorers for any chosen season.

No real statistics are pre-loaded — every Archive table ships with exactly
one clearly marked example row (fictitious team/player) to show the
expected format; delete it once you've pasted real data.
