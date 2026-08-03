import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

FONT = "Arial"
NAVY = "1F3864"
LIGHT_BLUE = "D9E2F3"
EXAMPLE_FILL = PatternFill("solid", fgColor="FFF2CC")
INPUT_FILL = PatternFill("solid", fgColor="FFFFCC")
HELPER_FILL = PatternFill("solid", fgColor="F2F2F2")
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
SUBHEADER_FILL = PatternFill("solid", fgColor=LIGHT_BLUE)
WHITE_BOLD = Font(name=FONT, bold=True, color="FFFFFF", size=11)
BOLD = Font(name=FONT, bold=True, size=11)
NORMAL = Font(name=FONT, size=11)
ITALIC_GREY = Font(name=FONT, italic=True, size=10, color="666666")
TITLE_FONT = Font(name=FONT, bold=True, size=16, color=NAVY)
SECTION_FONT = Font(name=FONT, bold=True, size=13, color=NAVY)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = Workbook()
wb.remove(wb.active)

def set_col_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

def style_all_arial(ws, max_row=200, max_col=40):
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            if cell.font is None or cell.font.name != FONT:
                if cell.font and (cell.font.bold or cell.font.size not in (None, 11)):
                    cell.font = Font(name=FONT, bold=cell.font.bold, size=cell.font.size or 11,
                                      color=cell.font.color, italic=cell.font.italic)
                else:
                    cell.font = Font(name=FONT, size=11)

def add_table(ws, name, ref, style="TableStyleMedium9"):
    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name=style, showFirstColumn=False,
                                         showLastColumn=False, showRowStripes=True,
                                         showColumnStripes=False)
    ws.add_table(tab)
    return tab

# ---------------------------------------------------------------
# LISTS (hidden helper sheet)
# ---------------------------------------------------------------
ws = wb.create_sheet("Lists")
ws["A1"] = "Season"
ws["A1"].font = BOLD
seasons = []
for start in range(2000, 2027):
    label = f"{start}/{str(start+1)[2:]}"
    seasons.append(label)
for i, label in enumerate(seasons, start=2):
    ws.cell(row=i, column=1, value=label)
ws.column_dimensions["A"].width = 12
style_all_arial(ws, max_row=30, max_col=2)
ws.sheet_state = "hidden"
N_SEASONS = len(seasons)  # 27
LAST_SEASON_ROW = 1 + N_SEASONS  # row 28

# ---------------------------------------------------------------
# README
# ---------------------------------------------------------------
ws = wb.create_sheet("README")
set_col_widths(ws, {"A": 3, "B": 105})
ws["B2"] = "Premiership Rugby Analytics Workbook"
ws["B2"].font = TITLE_FONT
ws["B3"] = "A standalone Excel tool for English Premiership Rugby team and player statistics, 2000/01 to present."
ws["B3"].font = Font(name=FONT, italic=True, size=11, color="444444")

lines = [
    ("How this workbook works", "section"),
    ("This workbook does NOT auto-scrape the web. Modern rugby stats sites (the official "
     "Premiership Rugby site, ESPN, statbunker, etc.) render their tables with JavaScript, "
     "which Excel's web tools cannot read automatically. Instead, this workbook uses a simple, "
     "reliable manual workflow: you visit a recommended page, copy the table you see on screen, "
     "paste it into a 'Raw Paste' sheet here, and built-in formulas instantly clean and reshape "
     "it into a permanent historical record. This is slower than a one-click refresh, but it "
     "works with ANY site, never breaks when a website redesigns itself, and needs no macros, "
     "add-ins, or external connections — fully standalone, as requested.", "body"),
    ("The 3-step workflow (repeat for each season / each update)", "section"),
    ("1. Go to the 'Sources' sheet and open one of the recommended pages for the stat type you "
     "want (team standings, core player stats, or advanced player stats).", "body"),
    ("2. Select and copy the table on that page, then paste it into the matching '...RawPaste' "
     "sheet at the marked paste zone. Fill in the Season and Source URL cells at the top.", "body"),
    ("3. Copy the 'Ready to Archive' block next to the paste zone (values only, excluding the "
     "last grey helper column) and paste it into the row directly below the last row of the "
     "matching '...Archive' table. Excel automatically extends the table and fills in the grey "
     "helper column for you.", "body"),
    ("Repeat this once per season to build up history back to 2000/01, then weekly during a "
     "season to keep it current.", "body"),
    ("What's in the workbook", "section"),
    ("• Sources — recommended pages to copy data from, and what to expect on each.", "bullet"),
    ("• TeamStats_RawPaste / TeamStats_Archive — core team standings, every season 2000/01+.", "bullet"),
    ("• PlayerCore_RawPaste / PlayerCore_Archive — core individual stats (appearances, tries, "
     "points, cards), every season 2000/01+.", "bullet"),
    ("• PlayerAdvanced_RawPaste / PlayerAdvanced_Archive — advanced match stats (carries, metres, "
     "defenders beaten, tackles, turnovers). Reliable public data for these generally only exists "
     "for roughly the last 10-15 seasons — kept as a separate module rather than forced into the "
     "full historic archive.", "bullet"),
    ("• Season_Explorer — pick any season from a dropdown and see its full standings table and "
     "chart, for season-to-season comparison.", "bullet"),
    ("• Dashboard — a team's league points trend across every season, and the top 10 try "
     "scorers for any chosen season.", "bullet"),
    ("• Setup_PowerQuery_Optional — an optional, ready-to-paste Power Query (M code) recipe "
     "that can auto-refresh the core team archive from Wikipedia, for anyone who wants to layer "
     "automation on top later. Entirely optional — the workbook is fully usable without it.", "bullet"),
    ("Important notes", "section"),
    ("• No real statistics are pre-loaded. Every Archive table ships with exactly one clearly "
     "marked EXAMPLE row using a fictitious team/player — delete it once you've pasted real data. "
     "This is deliberate: any numbers I typed in myself without being able to verify them live "
     "would risk being wrong, so the workbook starts empty and accurate rather than pre-filled "
     "and unverified.", "bullet"),
    ("• Recommended source pages (Sources sheet) were selected from general knowledge of these "
     "sites, not a live visit — always double check the column layout matches what you paste, "
     "and adjust the RawPaste header row if a site's columns differ.", "bullet"),
    ("• Team and competition names/sponsors have changed over the years (Allied Dunbar, Zurich, "
     "Guinness, Aviva, Gallagher Premiership are all the same competition). Use the season's own "
     "team names as shown on the source page for that season.", "bullet"),
]

r = 6
for text, kind in lines:
    cell = ws.cell(row=r, column=2, value=text)
    if kind == "section":
        cell.font = SECTION_FONT
        r += 1
    elif kind == "bullet":
        cell.font = NORMAL
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30
        r += 1
    else:
        cell.font = NORMAL
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 45
        r += 1
    r += 1

wb.move_sheet("README", offset=-len(wb.sheetnames))

print("README + Lists built")

# ---------------------------------------------------------------
# SOURCES
# ---------------------------------------------------------------
ws = wb.create_sheet("Sources")
set_col_widths(ws, {"A": 22, "B": 30, "C": 45, "D": 55})
ws["A1"] = "Recommended pages to copy data from"
ws["A1"].font = TITLE_FONT
ws.merge_cells("A1:D1")
ws["A2"] = ("I could not browse these live while building this workbook (this build environment "
            "has no internet access), so these are based on general knowledge of each site rather "
            "than a fresh visit. Check the columns you see on screen match the RawPaste header row "
            "before pasting - if they don't, just edit the header row to match what you copied.")
ws["A2"].font = ITALIC_GREY
ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A2:D2")
ws.row_dimensions[2].height = 45

headers = ["Data type", "Recommended site", "What to copy", "Notes"]
hr = 4
for i, h in enumerate(headers):
    c = ws.cell(row=hr, column=1 + i, value=h)
    c.font = WHITE_BOLD
    c.fill = HEADER_FILL
    c.alignment = Alignment(wrap_text=True, vertical="center")

rows_data = [
    ("Team standings (any season, 2000/01+)",
     "en.wikipedia.org - search \"20XX-YY Premiership Rugby\" (e.g. 2024-25 Premiership Rugby)",
     "The 'Final table' / league table (Pos, Team, Pld, W, D, L, PF, PA, Diff, BP, Pts)",
     "Most reliable long-term source - plain HTML tables, one page per season back to 2000/01. "
     "Older seasons may be titled by that year's sponsor (Allied Dunbar/Zurich/Guinness/Aviva "
     "Premiership) - Wikipedia usually redirects these, but search if a direct link 404s."),
    ("Team standings (current season, live)",
     "premiershiprugby.com - Tables page",
     "The live standings table",
     "Official source, good for the current season while it's in progress."),
    ("Core player stats (tries, points, appearances)",
     "premiershiprugby.com - Stats page, or Wikipedia season page's 'Individual statistics' "
     "section where present",
     "Leading try scorers / points scorers table",
     "Wikipedia often only lists the top ~10-20 for older seasons, not every player - that's "
     "expected and fine, paste what's shown."),
    ("Advanced player stats (carries, metres, tackles, defenders beaten, turnovers)",
     "premiershiprugby.com - Stats page (Attack / Defence / Discipline categories), or "
     "statbunker.com",
     "Whichever stat category table you need - one paste per category, repeated per season",
     "This level of detail is a modern tracking phenomenon - expect it to only be available for "
     "roughly the last 10-15 seasons, not back to 2000."),
    ("Fixtures / results",
     "premiershiprugby.com - Fixtures page, or the Wikipedia season page's results section",
     "The fixtures/results table",
     "Useful for the upcoming 2026/27 season once fixtures are announced."),
]
r = hr + 1
for row in rows_data:
    for i, val in enumerate(row):
        c = ws.cell(row=r, column=1 + i, value=val)
        c.font = NORMAL
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = BORDER
    ws.row_dimensions[r].height = 60
    r += 1

# ---------------------------------------------------------------
# Helper to build a RawPaste + Archive pair
# ---------------------------------------------------------------

def build_archive_sheet(name, table_name, headers, example_row, extra_cols=None):
    """headers: list of column headers (including any helper columns at the end).
    example_row: list of values matching headers, for the one demo row."""
    ws = wb.create_sheet(name)
    ncols = len(headers)
    for i, h in enumerate(headers, start=1):
        col = get_column_letter(i)
        ws.column_dimensions[col].width = 16
        c = ws.cell(row=1, column=i, value=h)
        c.font = WHITE_BOLD
        c.fill = HEADER_FILL
    for i, v in enumerate(example_row, start=1):
        c = ws.cell(row=2, column=i, value=v)
        c.font = NORMAL
        c.fill = EXAMPLE_FILL
    ref = f"A1:{get_column_letter(ncols)}2"
    add_table(ws, table_name, ref)
    ws.freeze_panes = "A2"
    cmt = Comment("EXAMPLE ROW using a fictitious team/player so nobody mistakes it for real "
                  "data. Delete this row once you've pasted real data below it.", "Workbook")
    ws["A2"].comment = cmt
    return ws

print("Sources built")

# ---------------------------------------------------------------
# TeamStats_Archive
# ---------------------------------------------------------------
team_headers = ["Season", "Team", "Position", "Played", "Won", "Drawn", "Lost",
                 "PointsFor", "PointsAgainst", "PointsDiff", "BonusPoints", "LeaguePoints",
                 "SourceURL", "DateAdded", "SeasonPositionKey"]
team_example = ["2025/26", "EXAMPLE RFC", 1, 22, 18, 1, 3, 650, 420, 230, 12, 84,
                 "https://en.wikipedia.org/wiki/2025-26_Premiership_Rugby", "=TODAY()",
                 '=A2&"|"&C2']
ws_team_arch = build_archive_sheet("TeamStats_Archive", "TeamStatsArchive", team_headers, team_example)
ws_team_arch.column_dimensions["B"].width = 22
ws_team_arch.column_dimensions["M"].width = 45
ws_team_arch.column_dimensions["O"].width = 20

# ---------------------------------------------------------------
# TeamStats_RawPaste
# ---------------------------------------------------------------
ws = wb.create_sheet("TeamStats_RawPaste")
set_col_widths(ws, {"A": 10, "B": 20, "C": 9, "D": 8, "E": 8, "F": 8, "G": 10, "H": 12,
                     "I": 11, "J": 12, "K": 12})
ws["A1"] = "TEAM STATS - RAW PASTE"
ws["A1"].font = TITLE_FONT
ws.merge_cells("A1:K1")
instructions = ("1) Open one of the recommended pages on the 'Sources' sheet.  2) Copy the full "
                "final league table.  3) Paste it starting at cell A7 below, keeping the header "
                "row intact (it's fine if your headers don't exactly match row 6 - just make sure "
                "the column ORDER matches: Position, Team, Played, Won, Drawn, Lost, Points For, "
                "Points Against, Points Diff, Bonus Points, League Points).  4) Fill in Season and "
                "Source URL below.  5) Copy the 'Ready to archive' block (columns N to AA only - "
                "NOT the grey Z column) and paste-values it into the row below the last row of "
                "the TeamStats_Archive table.")
ws["A2"] = instructions
ws["A2"].font = ITALIC_GREY
ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A2:K2")
ws.row_dimensions[2].height = 60

ws["A4"] = "Season:"
ws["A4"].font = BOLD
ws["B4"].fill = INPUT_FILL
ws["B4"].font = NORMAL
ws["A5"] = "Source URL:"
ws["A5"].font = BOLD
ws["B5"].fill = INPUT_FILL
ws["B5"].font = NORMAL
ws.merge_cells("B5:E5")

dv_season = DataValidation(type="list", formula1=f"=Lists!$A$2:$A${LAST_SEASON_ROW}", allow_blank=True)
ws.add_data_validation(dv_season)
dv_season.add(ws["B4"])

paste_headers = ["Position", "Team", "Played", "Won", "Drawn", "Lost", "PointsFor",
                  "PointsAgainst", "PointsDiff", "BonusPoints", "LeaguePoints"]
HR = 6
for i, h in enumerate(paste_headers, start=1):
    c = ws.cell(row=HR, column=i, value=h)
    c.font = WHITE_BOLD
    c.fill = HEADER_FILL

ws.cell(row=HR, column=14, value="Ready to archive → (copy N:AA only, paste-values below TeamStats_Archive table)")
ws.cell(row=HR, column=14).font = ITALIC_GREY
ws.merge_cells(start_row=HR, start_column=14, end_row=HR, end_column=28)

archive_block_headers = ["Season", "Team", "Position", "Played", "Won", "Drawn", "Lost",
                          "PointsFor", "PointsAgainst", "PointsDiff", "BonusPoints",
                          "LeaguePoints", "SourceURL", "DateAdded", "SeasonPositionKey"]
AR = HR + 1
for i, h in enumerate(archive_block_headers, start=14):
    c = ws.cell(row=AR, column=i, value=h)
    c.font = BOLD
    c.fill = SUBHEADER_FILL if i < 14 + len(archive_block_headers) - 1 else HELPER_FILL

FIRST_DATA_ROW = HR + 1
LAST_DATA_ROW = HR + 24  # room for up to 24 teams
for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
    for col in range(1, 12):
        ws.cell(row=row, column=col).font = NORMAL
    # N..AB formulas
    ws.cell(row=row, column=14, value=f'=IF($A{row}="","",$B$4)').font = NORMAL           # Season
    ws.cell(row=row, column=15, value=f'=IF($A{row}="","",TRIM($B{row}))').font = NORMAL  # Team
    ws.cell(row=row, column=16, value=f'=IF($A{row}="","",IFERROR(VALUE($A{row}),$A{row}))').font = NORMAL  # Position
    ws.cell(row=row, column=17, value=f'=IF($A{row}="","",IFERROR(VALUE($C{row}),$C{row}))').font = NORMAL  # Played
    ws.cell(row=row, column=18, value=f'=IF($A{row}="","",IFERROR(VALUE($D{row}),$D{row}))').font = NORMAL  # Won
    ws.cell(row=row, column=19, value=f'=IF($A{row}="","",IFERROR(VALUE($E{row}),$E{row}))').font = NORMAL  # Drawn
    ws.cell(row=row, column=20, value=f'=IF($A{row}="","",IFERROR(VALUE($F{row}),$F{row}))').font = NORMAL  # Lost
    ws.cell(row=row, column=21, value=f'=IF($A{row}="","",IFERROR(VALUE($G{row}),$G{row}))').font = NORMAL  # PointsFor
    ws.cell(row=row, column=22, value=f'=IF($A{row}="","",IFERROR(VALUE($H{row}),$H{row}))').font = NORMAL  # PointsAgainst
    ws.cell(row=row, column=23, value=f'=IF($A{row}="","",IFERROR(VALUE($I{row}),$I{row}))').font = NORMAL  # PointsDiff
    ws.cell(row=row, column=24, value=f'=IF($A{row}="","",IFERROR(VALUE($J{row}),$J{row}))').font = NORMAL  # BonusPoints
    ws.cell(row=row, column=25, value=f'=IF($A{row}="","",IFERROR(VALUE($K{row}),$K{row}))').font = NORMAL  # LeaguePoints
    ws.cell(row=row, column=26, value=f'=IF($A{row}="","",$B$5)').font = NORMAL              # SourceURL
    ws.cell(row=row, column=27, value=f'=IF($A{row}="","",$B$4)').font = NORMAL              # DateAdded placeholder (set below)
    ws.cell(row=row, column=28, value=f'=IF($A{row}="","",N{row}&"|"&P{row})').font = NORMAL # SeasonPositionKey
    ws.cell(row=row, column=28).fill = HELPER_FILL

# fix DateAdded to actually be TODAY(), not a duplicate of season
for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
    ws.cell(row=row, column=27, value=f'=IF($A{row}="","",TODAY())')

ws.freeze_panes = "A7"

print("TeamStats sheets built")

# ---------------------------------------------------------------
# PlayerCore_Archive  (Season, Player, Team, Appearances, Tries, Points, YellowCards,
#                      RedCards, SourceURL, DateAdded, TriesRankInSeason, SeasonTriesRankKey)
# ---------------------------------------------------------------
pc_headers = ["Season", "Player", "Team", "Appearances", "Tries", "Points", "YellowCards",
              "RedCards", "SourceURL", "DateAdded", "TriesRankInSeason", "SeasonTriesRankKey"]
pc_example = ["2025/26", "Example Player", "EXAMPLE RFC", 20, 15, 75, 1, 0,
              "https://en.wikipedia.org/wiki/2025-26_Premiership_Rugby", "=TODAY()", "", ""]
ws_pc = build_archive_sheet("PlayerCore_Archive", "PlayerCoreArchive", pc_headers, pc_example)
ws_pc.column_dimensions["B"].width = 20
ws_pc.column_dimensions["C"].width = 20
ws_pc.column_dimensions["I"].width = 45
ws_pc["K2"] = ('=SUMPRODUCT(($A$2:$A$5000=A2)*($E$2:$E$5000>E2))'
               '+SUMPRODUCT(($A$2:$A$5000=A2)*($E$2:$E$5000=E2)*(ROW($A$2:$A$5000)<ROW(A2)))+1')
ws_pc["L2"] = '=A2&"|"&K2'
ws_pc["K2"].fill = HELPER_FILL
ws_pc["L2"].fill = HELPER_FILL

ws = wb.create_sheet("PlayerCore_RawPaste")
set_col_widths(ws, {"A": 20, "B": 20, "C": 13, "D": 8, "E": 9, "F": 12, "G": 12})
ws["A1"] = "PLAYER CORE STATS - RAW PASTE"
ws["A1"].font = TITLE_FONT
ws.merge_cells("A1:G1")
instructions = ("1) Open a recommended page on 'Sources' for core player stats.  2) Copy the "
                "leading scorers / player stats table.  3) Paste it starting at cell A7, column "
                "order: Player, Team, Appearances, Tries, Points, Yellow Cards, Red Cards (leave "
                "cells blank if a source doesn't have a column - that's fine).  4) Fill in Season "
                "and Source URL below.  5) Copy the 'Ready to archive' block (columns J to Q only) "
                "and paste-values it into the row below the last row of the PlayerCore_Archive "
                "table.")
ws["A2"] = instructions
ws["A2"].font = ITALIC_GREY
ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A2:G2")
ws.row_dimensions[2].height = 60

ws["A4"] = "Season:"
ws["A4"].font = BOLD
ws["B4"].fill = INPUT_FILL
ws["B4"].font = NORMAL
ws["A5"] = "Source URL:"
ws["A5"].font = BOLD
ws["B5"].fill = INPUT_FILL
ws["B5"].font = NORMAL
ws.merge_cells("B5:E5")
dv = DataValidation(type="list", formula1=f"=Lists!$A$2:$A${LAST_SEASON_ROW}", allow_blank=True)
ws.add_data_validation(dv)
dv.add(ws["B4"])

paste_headers = ["Player", "Team", "Appearances", "Tries", "Points", "YellowCards", "RedCards"]
HR = 6
for i, h in enumerate(paste_headers, start=1):
    c = ws.cell(row=HR, column=i, value=h)
    c.font = WHITE_BOLD
    c.fill = HEADER_FILL

ws.cell(row=HR, column=10, value="Ready to archive → (copy J:Q only, paste-values below PlayerCore_Archive table)")
ws.cell(row=HR, column=10).font = ITALIC_GREY
ws.merge_cells(start_row=HR, start_column=10, end_row=HR, end_column=17)

archive_block_headers = ["Season", "Player", "Team", "Appearances", "Tries", "Points",
                          "YellowCards", "RedCards", "SourceURL", "DateAdded"]
AR = HR + 1
for i, h in enumerate(archive_block_headers, start=10):
    c = ws.cell(row=AR, column=i, value=h)
    c.font = BOLD
    c.fill = SUBHEADER_FILL

FIRST_DATA_ROW = HR + 1
LAST_DATA_ROW = HR + 60  # room for up to 60 players per paste
for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
    for col in range(1, 8):
        ws.cell(row=row, column=col).font = NORMAL
    ws.cell(row=row, column=10, value=f'=IF($A{row}="","",$B$4)')            # Season
    ws.cell(row=row, column=11, value=f'=IF($A{row}="","",TRIM($A{row}))')   # Player
    ws.cell(row=row, column=12, value=f'=IF($A{row}="","",TRIM($B{row}))')   # Team
    ws.cell(row=row, column=13, value=f'=IF($A{row}="","",IFERROR(VALUE($C{row}),$C{row}))')  # Appearances
    ws.cell(row=row, column=14, value=f'=IF($A{row}="","",IFERROR(VALUE($D{row}),$D{row}))')  # Tries
    ws.cell(row=row, column=15, value=f'=IF($A{row}="","",IFERROR(VALUE($E{row}),$E{row}))')  # Points
    ws.cell(row=row, column=16, value=f'=IF($A{row}="","",IFERROR(VALUE($F{row}),$F{row}))')  # YellowCards
    ws.cell(row=row, column=17, value=f'=IF($A{row}="","",IFERROR(VALUE($G{row}),$G{row}))')  # RedCards
    ws.cell(row=row, column=18, value=f'=IF($A{row}="","",$B$5)')            # SourceURL
    ws.cell(row=row, column=19, value=f'=IF($A{row}="","",TODAY())')         # DateAdded
    for col in (10, 11, 12, 13, 14, 15, 16, 17, 18, 19):
        ws.cell(row=row, column=col).font = NORMAL

ws.freeze_panes = "A7"
print("PlayerCore sheets built")

# ---------------------------------------------------------------
# PlayerAdvanced_Archive
# ---------------------------------------------------------------
pa_headers = ["Season", "Player", "Team", "Carries", "MetresMade", "DefendersBeaten",
              "CleanBreaks", "Offloads", "Tackles", "TackleSuccessPct", "Turnovers",
              "SourceURL", "DateAdded"]
pa_example = ["2025/26", "Example Player", "EXAMPLE RFC", 110, 480, 22, 8, 14, 95, 0.87, 6,
              "https://www.premiershiprugby.com/stats/", "=TODAY()"]
ws_pa = build_archive_sheet("PlayerAdvanced_Archive", "PlayerAdvancedArchive", pa_headers, pa_example)
ws_pa.column_dimensions["B"].width = 20
ws_pa.column_dimensions["C"].width = 20
ws_pa.column_dimensions["L"].width = 45
ws_pa["J2"].number_format = "0.0%"

ws = wb.create_sheet("PlayerAdvanced_RawPaste")
set_col_widths(ws, {"A": 20, "B": 20, "C": 9, "D": 11, "E": 15, "F": 12, "G": 10, "H": 9, "I": 14, "J": 10})
ws["A1"] = "PLAYER ADVANCED STATS - RAW PASTE"
ws["A1"].font = TITLE_FONT
ws.merge_cells("A1:J1")
instructions = ("This module is for modern match stats (carries, metres, defenders beaten, "
                 "tackles, turnovers) - reliable data for these typically only exists for roughly "
                 "the last 10-15 seasons, so don't worry about backfilling to 2000.  1) Open a "
                 "recommended page on 'Sources' for advanced player stats (pick one stat category "
                 "at a time, e.g. Attack, Defence).  2) Copy that table.  3) Paste it starting at "
                 "A7, column order: Player, Team, Carries, Metres Made, Defenders Beaten, Clean "
                 "Breaks, Offloads, Tackles, Tackle Success %, Turnovers - leave blank any columns "
                 "your source doesn't have.  4) Fill in Season and Source URL below.  5) Copy the "
                 "'Ready to archive' block (columns M to Y) and paste-values it below the last row "
                 "of PlayerAdvanced_Archive.")
ws["A2"] = instructions
ws["A2"].font = ITALIC_GREY
ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A2:J2")
ws.row_dimensions[2].height = 75

ws["A4"] = "Season:"
ws["A4"].font = BOLD
ws["B4"].fill = INPUT_FILL
ws["B4"].font = NORMAL
ws["A5"] = "Source URL:"
ws["A5"].font = BOLD
ws["B5"].fill = INPUT_FILL
ws["B5"].font = NORMAL
ws.merge_cells("B5:E5")
dv = DataValidation(type="list", formula1=f"=Lists!$A$2:$A${LAST_SEASON_ROW}", allow_blank=True)
ws.add_data_validation(dv)
dv.add(ws["B4"])

paste_headers = ["Player", "Team", "Carries", "MetresMade", "DefendersBeaten", "CleanBreaks",
                  "Offloads", "Tackles", "TackleSuccessPct", "Turnovers"]
HR = 6
for i, h in enumerate(paste_headers, start=1):
    c = ws.cell(row=HR, column=i, value=h)
    c.font = WHITE_BOLD
    c.fill = HEADER_FILL

ws.cell(row=HR, column=13, value="Ready to archive → (copy M:Y only, paste-values below PlayerAdvanced_Archive table)")
ws.cell(row=HR, column=13).font = ITALIC_GREY
ws.merge_cells(start_row=HR, start_column=13, end_row=HR, end_column=25)

archive_block_headers = ["Season", "Player", "Team", "Carries", "MetresMade", "DefendersBeaten",
                          "CleanBreaks", "Offloads", "Tackles", "TackleSuccessPct", "Turnovers",
                          "SourceURL", "DateAdded"]
AR = HR + 1
for i, h in enumerate(archive_block_headers, start=13):
    c = ws.cell(row=AR, column=i, value=h)
    c.font = BOLD
    c.fill = SUBHEADER_FILL

FIRST_DATA_ROW = HR + 1
LAST_DATA_ROW = HR + 60
for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
    for col in range(1, 11):
        ws.cell(row=row, column=col).font = NORMAL
    ws.cell(row=row, column=13, value=f'=IF($A{row}="","",$B$4)')            # Season
    ws.cell(row=row, column=14, value=f'=IF($A{row}="","",TRIM($A{row}))')   # Player
    ws.cell(row=row, column=15, value=f'=IF($A{row}="","",TRIM($B{row}))')   # Team
    ws.cell(row=row, column=16, value=f'=IF($A{row}="","",IFERROR(VALUE($C{row}),$C{row}))')  # Carries
    ws.cell(row=row, column=17, value=f'=IF($A{row}="","",IFERROR(VALUE($D{row}),$D{row}))')  # MetresMade
    ws.cell(row=row, column=18, value=f'=IF($A{row}="","",IFERROR(VALUE($E{row}),$E{row}))')  # DefendersBeaten
    ws.cell(row=row, column=19, value=f'=IF($A{row}="","",IFERROR(VALUE($F{row}),$F{row}))')  # CleanBreaks
    ws.cell(row=row, column=20, value=f'=IF($A{row}="","",IFERROR(VALUE($G{row}),$G{row}))')  # Offloads
    ws.cell(row=row, column=21, value=f'=IF($A{row}="","",IFERROR(VALUE($H{row}),$H{row}))')  # Tackles
    ws.cell(row=row, column=22, value=f'=IF($A{row}="","",IFERROR(VALUE($I{row}),$I{row}))')  # TackleSuccessPct
    ws.cell(row=row, column=23, value=f'=IF($A{row}="","",IFERROR(VALUE($J{row}),$J{row}))')  # Turnovers
    ws.cell(row=row, column=24, value=f'=IF($A{row}="","",$B$5)')            # SourceURL
    ws.cell(row=row, column=25, value=f'=IF($A{row}="","",TODAY())')         # DateAdded
    for col in range(13, 26):
        ws.cell(row=row, column=col).font = NORMAL

ws.freeze_panes = "A7"
print("PlayerAdvanced sheets built")

# ---------------------------------------------------------------
# Season_Explorer
# ---------------------------------------------------------------
ws = wb.create_sheet("Season_Explorer")
set_col_widths(ws, {"A": 10, "B": 22, "C": 9, "D": 8, "E": 8, "F": 8, "G": 11, "H": 13,
                     "I": 11, "J": 11, "K": 12, "M": 20})
ws["A1"] = "Season Explorer"
ws["A1"].font = TITLE_FONT
ws["A2"] = "Pick a season to see its full standings table below, for season-to-season comparison."
ws["A2"].font = ITALIC_GREY
ws["B4"] = "Select Season:"
ws["B4"].font = BOLD
ws["C4"].fill = INPUT_FILL
ws["C4"].font = BOLD
dv = DataValidation(type="list", formula1=f"=Lists!$A$2:$A${LAST_SEASON_ROW}", allow_blank=True)
ws.add_data_validation(dv)
dv.add(ws["C4"])
ws["C4"] = seasons[-1]  # default to most recent season

se_headers = ["Position", "Team", "Played", "Won", "Drawn", "Lost", "PointsFor",
              "PointsAgainst", "PointsDiff", "BonusPoints", "LeaguePoints"]
HR = 6
for i, h in enumerate(se_headers, start=1):
    c = ws.cell(row=HR, column=i, value=h)
    c.font = WHITE_BOLD
    c.fill = HEADER_FILL
ws.cell(row=HR, column=13, value="Key")
ws.cell(row=HR, column=13).font = ITALIC_GREY

N_ROWS = 16  # supports up to 16 teams in a season
FIRST = HR + 1
LAST = HR + N_ROWS
for k, row in enumerate(range(FIRST, LAST + 1), start=1):
    ws.cell(row=row, column=13, value=f'=$C$4&"|"&{k}')
    ws.cell(row=row, column=13).font = NORMAL
    ws.cell(row=row, column=13).fill = HELPER_FILL
    key_ref = f"$M{row}"
    ws.cell(row=row, column=1, value=f'=IFERROR(INDEX(TeamStatsArchive[Position],MATCH({key_ref},TeamStatsArchive[SeasonPositionKey],0)),"")')
    ws.cell(row=row, column=2, value=f'=IFERROR(INDEX(TeamStatsArchive[Team],MATCH({key_ref},TeamStatsArchive[SeasonPositionKey],0)),"")')
    ws.cell(row=row, column=3, value=f'=IFERROR(INDEX(TeamStatsArchive[Played],MATCH({key_ref},TeamStatsArchive[SeasonPositionKey],0)),"")')
    ws.cell(row=row, column=4, value=f'=IFERROR(INDEX(TeamStatsArchive[Won],MATCH({key_ref},TeamStatsArchive[SeasonPositionKey],0)),"")')
    ws.cell(row=row, column=5, value=f'=IFERROR(INDEX(TeamStatsArchive[Drawn],MATCH({key_ref},TeamStatsArchive[SeasonPositionKey],0)),"")')
    ws.cell(row=row, column=6, value=f'=IFERROR(INDEX(TeamStatsArchive[Lost],MATCH({key_ref},TeamStatsArchive[SeasonPositionKey],0)),"")')
    ws.cell(row=row, column=7, value=f'=IFERROR(INDEX(TeamStatsArchive[PointsFor],MATCH({key_ref},TeamStatsArchive[SeasonPositionKey],0)),"")')
    ws.cell(row=row, column=8, value=f'=IFERROR(INDEX(TeamStatsArchive[PointsAgainst],MATCH({key_ref},TeamStatsArchive[SeasonPositionKey],0)),"")')
    ws.cell(row=row, column=9, value=f'=IFERROR(INDEX(TeamStatsArchive[PointsDiff],MATCH({key_ref},TeamStatsArchive[SeasonPositionKey],0)),"")')
    ws.cell(row=row, column=10, value=f'=IFERROR(INDEX(TeamStatsArchive[BonusPoints],MATCH({key_ref},TeamStatsArchive[SeasonPositionKey],0)),"")')
    ws.cell(row=row, column=11, value=f'=IFERROR(INDEX(TeamStatsArchive[LeaguePoints],MATCH({key_ref},TeamStatsArchive[SeasonPositionKey],0)),"")')
    for col in range(1, 12):
        ws.cell(row=row, column=col).font = NORMAL
        ws.cell(row=row, column=col).border = BORDER

chart = BarChart()
chart.title = "Standings - League Points by Team (selected season)"
chart.y_axis.title = "League Points"
chart.x_axis.title = "Team"
chart.style = 10
data = Reference(ws, min_col=11, min_row=HR, max_row=LAST)
cats = Reference(ws, min_col=2, min_row=FIRST, max_row=LAST)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.width = 22
chart.height = 11
ws.add_chart(chart, "B22")

print("Season_Explorer built")

# ---------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------
ws = wb.create_sheet("Dashboard")
set_col_widths(ws, {"A": 14, "B": 22, "C": 14, "E": 6, "F": 20, "G": 22, "H": 8})
ws["A1"] = "Dashboard"
ws["A1"].font = TITLE_FONT

# --- Section 1: team trend across seasons ---
ws["A3"] = "Team trend across seasons"
ws["A3"].font = SECTION_FONT
ws["A4"] = "Select Team (type exactly as it appears in TeamStats_Archive):"
ws["A4"].font = NORMAL
ws["A4"].alignment = Alignment(wrap_text=True)
ws.merge_cells("A4:B4")
ws["C4"].fill = INPUT_FILL
ws["C4"].font = BOLD
ws["C4"] = "EXAMPLE RFC"

ws["A6"] = "Season"
ws["B6"] = "LeaguePoints"
ws["A6"].font = WHITE_BOLD
ws["B6"].font = WHITE_BOLD
ws["A6"].fill = HEADER_FILL
ws["B6"].fill = HEADER_FILL
for i, label in enumerate(seasons, start=7):
    ws.cell(row=i, column=1, value=label).font = NORMAL
    ws.cell(row=i, column=2,
            value=f'=IFERROR(SUMIFS(TeamStatsArchive[LeaguePoints],TeamStatsArchive[Season],A{i},TeamStatsArchive[Team],$C$4),"")').font = NORMAL
LAST_TREND_ROW = 6 + len(seasons)

line = LineChart()
line.title = "League Points by Season - selected team"
line.y_axis.title = "League Points"
line.x_axis.title = "Season"
data = Reference(ws, min_col=2, min_row=6, max_row=LAST_TREND_ROW)
cats = Reference(ws, min_col=1, min_row=7, max_row=LAST_TREND_ROW)
line.add_data(data, titles_from_data=True)
line.set_categories(cats)
line.width = 22
line.height = 10
ws.add_chart(line, "D3")

# --- Section 2: top try scorers for a chosen season ---
top_row0 = LAST_TREND_ROW + 3
ws.cell(row=top_row0, column=1, value="Top 10 try scorers - selected season").font = SECTION_FONT
ws.cell(row=top_row0 + 1, column=1, value="Select Season:").font = BOLD
ws.cell(row=top_row0 + 1, column=2).fill = INPUT_FILL
ws.cell(row=top_row0 + 1, column=2).font = BOLD
dv2 = DataValidation(type="list", formula1=f"=Lists!$A$2:$A${LAST_SEASON_ROW}", allow_blank=True)
ws.add_data_validation(dv2)
dv2.add(ws.cell(row=top_row0 + 1, column=2))
ws.cell(row=top_row0 + 1, column=2, value=seasons[-1])

hdr_row = top_row0 + 3
for i, h in enumerate(["Rank", "Player", "Team", "Tries", "Key"], start=1):
    c = ws.cell(row=hdr_row, column=i, value=h)
    c.font = WHITE_BOLD
    c.fill = HEADER_FILL
season_cell = f"$B${top_row0 + 1}"
first_top = hdr_row + 1
last_top = hdr_row + 10
for k, row in enumerate(range(first_top, last_top + 1), start=1):
    ws.cell(row=row, column=1, value=k).font = NORMAL
    ws.cell(row=row, column=5, value=f'={season_cell}&"|"&A{row}')
    ws.cell(row=row, column=5).font = NORMAL
    ws.cell(row=row, column=5).fill = HELPER_FILL
    key_ref = f"$E{row}"
    ws.cell(row=row, column=2, value=f'=IFERROR(INDEX(PlayerCoreArchive[Player],MATCH({key_ref},PlayerCoreArchive[SeasonTriesRankKey],0)),"")').font = NORMAL
    ws.cell(row=row, column=3, value=f'=IFERROR(INDEX(PlayerCoreArchive[Team],MATCH({key_ref},PlayerCoreArchive[SeasonTriesRankKey],0)),"")').font = NORMAL
    ws.cell(row=row, column=4, value=f'=IFERROR(INDEX(PlayerCoreArchive[Tries],MATCH({key_ref},PlayerCoreArchive[SeasonTriesRankKey],0)),"")').font = NORMAL
    for col in range(1, 5):
        ws.cell(row=row, column=col).border = BORDER

bar = BarChart()
bar.type = "bar"
bar.title = "Top 10 try scorers - selected season"
bar.y_axis.title = "Player"
bar.x_axis.title = "Tries"
data = Reference(ws, min_col=4, min_row=hdr_row, max_row=last_top)
cats = Reference(ws, min_col=2, min_row=first_top, max_row=last_top)
bar.add_data(data, titles_from_data=True)
bar.set_categories(cats)
bar.width = 22
bar.height = 11
ws.add_chart(bar, f"G{top_row0 + 1}")

print("Dashboard built")

# ---------------------------------------------------------------
# Setup_PowerQuery_Optional
# ---------------------------------------------------------------
ws = wb.create_sheet("Setup_PowerQuery_Optional")
set_col_widths(ws, {"A": 100})
ws["A1"] = "Optional: Power Query auto-refresh recipe (Wikipedia team standings)"
ws["A1"].font = TITLE_FONT
ws.merge_cells("A1:A1")

pq_text = [
    ("This is entirely optional. The workbook works fully without it - this is for anyone who "
     "wants to layer one-click automation on top of the manual workflow, for the Team Stats "
     "archive specifically (Wikipedia's plain HTML tables are the one source stable and simple "
     "enough for Power Query to read reliably; the player-stats sites generally are not, for the "
     "JavaScript-rendering reasons explained on the README sheet).", "body"),
    ("This M code was written from general knowledge of Wikipedia's page structure, not tested "
     "live against the current pages (this workbook was built in an environment without internet "
     "access) - the article titles for older seasons in particular may need adjusting; expect to "
     "debug it a little the first time you run it.", "body"),
    ("How to set it up:", "section"),
    ("1. In Excel, go to Data > Get Data > From Other Sources > Blank Query.", "body"),
    ("2. In the query editor, go to Home > Advanced Editor, delete the placeholder text, and "
     "paste in the M code below.", "body"),
    ("3. Click Done, then Close & Load To... and load it as a new table (or load it into the "
     "TeamStats_Archive table's query, if you're comfortable rewiring it).", "body"),
    ("4. Right-click the query and set it to refresh automatically, or use Data > Refresh All.", "body"),
    ("M code:", "section"),
]
r = 3
for text, kind in pq_text:
    cell = ws.cell(row=r, column=1, value=text)
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    if kind == "section":
        cell.font = SECTION_FONT
    else:
        cell.font = NORMAL
        ws.row_dimensions[r].height = 30
    r += 1

m_code = '''let
    Seasons = {2000..2026},
    GetSeasonTable = (StartYear as number) =>
        let
            EndYearShort = Text.End(Number.ToText(StartYear + 1), 2),
            SeasonLabel = Number.ToText(StartYear) & "/" & EndYearShort,
            PageTitle = Number.ToText(StartYear) & "-" & EndYearShort & "_Premiership_Rugby",
            Url = "https://en.wikipedia.org/wiki/" & PageTitle,
            Result =
                try
                    let
                        Source = Web.Page(Web.Contents(Url)),
                        // Pick out the table that looks like the final league table -
                        // adjust the index [0] below if the wrong table comes back.
                        Candidates = Table.SelectRows(Source, each Table.ColumnCount(Data) > 5),
                        LeagueTable = Candidates{0}[Data],
                        Promoted = Table.PromoteHeaders(LeagueTable, [PromoteAllScalars=true]),
                        WithSeason = Table.AddColumn(Promoted, "Season", each SeasonLabel)
                    in
                        WithSeason
                otherwise
                    null
        in
            Result,
    Tables = List.Transform(Seasons, each GetSeasonTable(_)),
    NonNull = List.Select(Tables, each _ <> null),
    Combined = Table.Combine(NonNull)
in
    Combined'''

ws.cell(row=r, column=1, value=m_code)
ws.cell(row=r, column=1).font = Font(name="Consolas", size=10)
ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[r].height = 420

print("Setup_PowerQuery_Optional built")

# ---------------------------------------------------------------
# Finalize: sheet order, active sheet, save
# ---------------------------------------------------------------
order = ["README", "Sources", "TeamStats_RawPaste", "TeamStats_Archive",
         "PlayerCore_RawPaste", "PlayerCore_Archive",
         "PlayerAdvanced_RawPaste", "PlayerAdvanced_Archive",
         "Season_Explorer", "Dashboard", "Setup_PowerQuery_Optional", "Lists"]
for name in order:
    wb.move_sheet(name, offset=len(wb.sheetnames))
wb.active = 0

OUT = "/tmp/claude-0/-home-user-applescripts/7e4ce647-099b-51f7-9c5d-57b91fc3d77c/scratchpad/Premiership_Rugby_Analytics.xlsx"
wb.save(OUT)
print("Saved:", OUT)
