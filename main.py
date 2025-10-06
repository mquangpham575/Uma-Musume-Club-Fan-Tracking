import asyncio
import json
import os
import sys
from pathlib import Path
import pandas as pd
import zendriver as zd
import gspread
from google.oauth2.service_account import Credentials

from globals import CLUBS  # external config for club list/URLs/thresholds


# ========== Google Sheets config ==========
SHEET_ID = "1dA2gLLQY5RA23gWFunytA50xXOQ5oPgKb1TjqVCTNlk"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CREDS = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
GC = gspread.authorize(CREDS)


# === Club selection ===
def pick_club() -> dict | str:
    print("=== Choose a club to export ===")
    for key, cfg in CLUBS.items():
        print(f"{key}. {cfg['title']}")
    print("0. Export ALL clubs")
    choice = input("Enter 0–7: ").strip()
    if choice == "0":
        return "ALL"
    if choice not in CLUBS:
        print("Invalid choice, defaulting to 1.")
        choice = "1"
    return CLUBS[choice]


# === Paths ===
def resolve_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


# === Data fetch ===
async def fetch_json(URL: str):
    browser = await zd.start(
        browser="edge",
        browser_executable_path="C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
    )
    page = await browser.get("https://google.com")
    async with page.expect_request(r".*/api/.*") as req:
        await page.get(URL)
        await req.value
        body, _ = await req.response_body
    await browser.stop()
    text = body.decode("utf-8", errors="replace") if isinstance(body, (bytes, bytearray)) else str(body)
    return json.loads(text)


# === DataFrame processing ===
def build_dataframe(data: dict) -> pd.DataFrame:
    df = pd.json_normalize(data.get("club_friend_history") or [])
    for c in ("friend_viewer_id", "friend_name", "actual_date", "adjusted_interpolated_fan_gain"):
        if c not in df.columns:
            df[c] = pd.NA

    df = (
        df.assign(day_col=lambda d: "Day " + d["actual_date"].astype(str))
          .pivot_table(
              index=["friend_viewer_id", "friend_name"],
              columns="day_col",
              values="adjusted_interpolated_fan_gain",
              aggfunc="first"
          )
          .reset_index()
    )
    df.columns.name = None

    def _day_key(x: str):
        if not isinstance(x, str) or not x.startswith("Day "): return x
        try:
            return int(x.split(maxsplit=1)[1])
        except Exception:
            return x

    day_cols = sorted([c for c in df.columns if isinstance(c, str) and c.startswith("Day ")], key=_day_key)

    df["AVG/d"] = df[day_cols].mean(axis=1).round(0) if day_cols else 0
    df = df[["friend_viewer_id", "friend_name", "AVG/d"] + day_cols].rename(
        columns={"friend_viewer_id": "Member_ID", "friend_name": "Member_Name"}
    )
    df["Member_ID"] = df["Member_ID"].fillna("").astype(str)
    df["Member_Name"] = df["Member_Name"].fillna("").astype(str)
    for c in df.columns:
        if c not in ("Member_ID", "Member_Name"):
            df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df.sort_values(["AVG/d", "Member_Name"], ascending=[False, True], kind="mergesort").reset_index(drop=True)
    return df


# === Google Sheets export ===
def export_to_gsheets(df: pd.DataFrame, spreadsheet_id: str, sheet_title: str, threshold: int):
    from gspread.utils import rowcol_to_a1

    GAP_COL = " "
    dcols = [c for c in df.columns if isinstance(c, str) and c.startswith("Day ")]
    df_to_write = df.copy()

    if dcols:
        df_to_write["Total"] = df_to_write[dcols].sum(axis=1, min_count=1)
        gidx = df_to_write.columns.get_loc("Total")
        df_to_write.insert(gidx, GAP_COL, "")
    else:
        gidx = None

    bottom_totals = {}
    for c in df_to_write.columns:
        if c == "Member_Name":
            bottom_totals[c] = "Total"
        elif c in ("Member_ID", GAP_COL):
            bottom_totals[c] = ""
        else:
            bottom_totals[c] = pd.to_numeric(df_to_write[c], errors="coerce").sum(min_count=1)

    header = list(map(str, df_to_write.columns))
    data_rows = df_to_write.where(pd.notna(df_to_write), "").values.tolist()
    totals_row = [("" if pd.isna(v) else v) for v in (bottom_totals.get(c, "") for c in df_to_write.columns)]
    values = [header] + data_rows + [totals_row]

    ss = GC.open_by_key(spreadsheet_id)
    for ws in ss.worksheets():
        if ws.title == sheet_title:
            ss.del_worksheet(ws)
            break
    ws = ss.add_worksheet(title=sheet_title, rows=max(len(values) + 50, 100), cols=max(len(header) + 10, 26))

    end_row = len(values)
    end_col = len(header)
    end_a1 = rowcol_to_a1(end_row, end_col)
    ws.update(values, f"A1:{end_a1}")

    # ===== FORMATTING =====
    sheet_id = ws._properties["sheetId"]
    last_data_row_1based = 1 + len(data_rows)

    header_range = {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": end_col}
    totals_range = {"sheetId": sheet_id, "startRowIndex": end_row - 1, "endRowIndex": end_row, "startColumnIndex": 0, "endColumnIndex": end_col}
    blank_col_range = {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": end_row, "startColumnIndex": gidx, "endColumnIndex": gidx + 1} if gidx is not None else None
    header_plus_data_range = {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": last_data_row_1based, "startColumnIndex": 0, "endColumnIndex": end_col}
    data_rows_range = {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": last_data_row_1based, "startColumnIndex": 0, "endColumnIndex": end_col}
    full_table_range = {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": end_row, "startColumnIndex": 0, "endColumnIndex": end_col}

    skip = {"Member_ID", "Member_Name", GAP_COL}
    numeric_cols_1 = [i + 1 for i, c in enumerate(header) if c not in skip]

    def col_data_grid(col_1):
        return {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": last_data_row_1based, "startColumnIndex": col_1 - 1, "endColumnIndex": col_1}
    numeric_ranges = [col_data_grid(c1) for c1 in numeric_cols_1]

    blue_fill = {"red": 0.31, "green": 0.51, "blue": 0.74}
    white_font = {"red": 1, "green": 1, "blue": 1}
    red_fill   = {"red": 1.00, "green": 0.78, "blue": 0.81}
    grey_fill  = {"red": 0.75, "green": 0.75, "blue": 0.75}

    requests = [
        {"setBasicFilter": {"filter": {"range": header_plus_data_range}}},
        {
            "repeatCell": {
                "range": header_range,
                "cell": {"userEnteredFormat": {"backgroundColor": blue_fill, "textFormat": {"bold": True, "foregroundColor": white_font}}},
                "fields": "userEnteredFormat(backgroundColor,textFormat)"
            }
        },
        {
            "repeatCell": {
                "range": totals_range,
                "cell": {"userEnteredFormat": {"backgroundColor": blue_fill, "textFormat": {"bold": True, "foregroundColor": white_font}}},
                "fields": "userEnteredFormat(backgroundColor,textFormat)"
            }
        },
        *([
            {
                "repeatCell": {
                    "range": blank_col_range,
                    "cell": {"userEnteredFormat": {"backgroundColor": blue_fill}},
                    "fields": "userEnteredFormat.backgroundColor"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": gidx, "endIndex": gidx + 1},
                    "properties": {"pixelSize": 40},
                    "fields": "pixelSize"
                }
            }
        ] if blank_col_range else []),
        {
            "addConditionalFormatRule": {
                "rule": {"ranges": numeric_ranges, "booleanRule": {"condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": str(threshold)}]}, "format": {"backgroundColor": red_fill}}},
                "index": 0
            }
        },
        {
            "addConditionalFormatRule": {
                "rule": {"ranges": numeric_ranges, "booleanRule": {"condition": {"type": "BLANK"}, "format": {"backgroundColor": grey_fill}}},
                "index": 0
            }
        },
        {
            "updateBorders": {
                "range": full_table_range,
                "top": {"style": "SOLID"},
                "bottom": {"style": "SOLID"},
                "left": {"style": "SOLID"},
                "right": {"style": "SOLID"},
                "innerHorizontal": {"style": "SOLID"},
                "innerVertical": {"style": "SOLID"},
            }
        },
    ]

    ws.spreadsheet.batch_update({"requests": requests})
    print(f"✅ Exported '{sheet_title}' ({end_row} rows × {end_col} cols)")


# === Main ===
async def main():
    choice = pick_club()

    if choice == "ALL":
        print("\nExporting ALL clubs...\n")
        for key, cfg in CLUBS.items():
            print(f"→ Exporting {cfg['title']} ...")
            data = await fetch_json(cfg["URL"])
            df = build_dataframe(data)
            export_to_gsheets(df, spreadsheet_id=SHEET_ID, sheet_title=cfg["title"], threshold=cfg["THRESHOLD"])
        print("\n✅ All clubs exported successfully!")
    else:
        cfg = choice
        print(f"\nSelected: {cfg['title']}\nURL: {cfg['URL']}\nSheet: {SHEET_ID}\nThreshold: {cfg['THRESHOLD']}\n")
        data = await fetch_json(cfg["URL"])
        df = build_dataframe(data)
        export_to_gsheets(df, spreadsheet_id=SHEET_ID, sheet_title=cfg["title"], threshold=cfg["THRESHOLD"])
        print(f"✅ Exported single club '{cfg['title']}' successfully!")


if __name__ == "__main__":
    asyncio.run(main())
