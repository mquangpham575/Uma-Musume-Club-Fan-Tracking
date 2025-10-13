import asyncio
import json
import os
import sys
from pathlib import Path
import pandas as pd
import zendriver as zd
import gspread
from google.oauth2.service_account import Credentials
import time

from globals import CLUBS, SHEET_ID


# ========== Google Sheets config ==========
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CREDS = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
GC = gspread.authorize(CREDS)


# === Club selection ===
def pick_club() -> dict | str:
    print("=== Choose a club to export ===")
    for key, cfg in CLUBS.items():
        print(f"{key}. {cfg['title']}")
    print("0. Export ALL clubs (default)")
    choice = input("Enter 0–7 [default=0]: ").strip()
    if choice == "" or choice == "0":
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
    MAX_RETRIES = 3
    RETRY_DELAY = 5
    
    for attempt in range(MAX_RETRIES):
        browser = None
        try:
            browser = await zd.start(
                browser="edge",
                browser_executable_path="C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
            )
            page = await browser.get("https://google.com")
            
            async with page.expect_request(r".*\/api\/club_profile.*") as req:
                await page.get(URL)
                await req.value
                body, _ = await req.response_body
            
            text = body.decode("utf-8", errors="replace") if isinstance(body, (bytes, bytearray)) else str(body)
            return json.loads(text)
            
        except (zd.errors.RemoteDisconnectedError, zd.errors.ConnectionAbortedError) as e:
            print(f"Lỗi kết nối ({URL}, lần {attempt + 1}/{MAX_RETRIES}): {type(e).__name__}. Đang thử lại sau {RETRY_DELAY}s...")
            if attempt < MAX_RETRIES - 1:
                if browser:
                    await browser.stop()
                await asyncio.sleep(RETRY_DELAY)
                continue
            else:
                raise
                
        except Exception as e:
            raise e
            
        finally:
            if browser:
                await browser.stop()
    
    raise Exception(f"Thất bại sau {MAX_RETRIES} lần thử.")


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

    def _day_num(x: str):
        if not isinstance(x, str) or not x.startswith("Day "):
            return None
        try:
            return int(x.split(maxsplit=1)[1])
        except Exception:
            return None

    day_cols = [c for c in df.columns if isinstance(c, str) and c.startswith("Day ")]

    # --- Keep only members who have value on the newest day (max Day N) ---
    nums = [n for n in map(_day_num, day_cols) if n is not None]
    if nums:
        latest_day = max(nums)
        latest_col = f"Day {latest_day}"
        if latest_col in df.columns:
            df = df[~df[latest_col].isna()].copy()

    # Order Day columns numerically
    day_cols = sorted(day_cols, key=lambda c: (_day_num(c) if _day_num(c) is not None else float("inf")))

    # Compute AVG/d and finalize columns
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

    # ====== PREP DATA ======
    GAP_COL = " "
    dcols = [c for c in df.columns if isinstance(c, str) and c.startswith("Day ")]
    df_to_write = df.copy()

    # Add Total and a blue gap column before it
    if dcols:
        df_to_write["Total"] = df_to_write[dcols].sum(axis=1, min_count=1)
        gidx = df_to_write.columns.get_loc("Total")
        df_to_write.insert(gidx, GAP_COL, "")
    else:
        gidx = None

    # Bottom "Total" row (sum)
    bottom_totals = {}
    for c in df_to_write.columns:
        if c == "Member_Name":
            bottom_totals[c] = "Total"
        elif c in ("Member_ID", GAP_COL):
            bottom_totals[c] = ""
        else:
            bottom_totals[c] = pd.to_numeric(df_to_write[c], errors="coerce").sum(min_count=1)

    # Day AVG row — per-day means only (no AVG/d)
    day_avgs = pd.Series("", index=df_to_write.columns, dtype=object)
    if dcols:
        means = df_to_write[dcols].mean(axis=0, skipna=True).round(0)
        for c in dcols:
            day_avgs[c] = means.get(c, "")
    day_avgs["Member_Name"] = "Day AVG"

    header = list(map(str, df_to_write.columns))
    data_rows = df_to_write.where(pd.notna(df_to_write), "").values.tolist()
    totals_row = [("" if pd.isna(v) else v) for v in (bottom_totals.get(c, "") for c in df_to_write.columns)]
    day_avg_row = [day_avgs.get(c, "") for c in df_to_write.columns]

    # Values order: header, data..., Total, Day AVG
    values = [header] + data_rows + [totals_row, day_avg_row]

    # ====== OPEN SHEET ======
    ss = GC.open_by_key(spreadsheet_id)
    for ws in ss.worksheets():
        if ws.title == sheet_title:
            ss.del_worksheet(ws)
            break
    ws = ss.add_worksheet(title=sheet_title, rows=max(len(values) + 50, 120), cols=max(len(header) + 10, 26))

    # Write values
    end_row = len(values)
    end_col = len(header)
    end_a1 = rowcol_to_a1(end_row, end_col)
    ws.update(values, f"A1:{end_a1}")

    # ====== FORMATTING ======
    sheet_id = ws._properties["sheetId"]
    last_data_row_1based = 1 + len(data_rows)  # header + data (excludes the 2 summary rows)

    header_range = {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": end_col}
    totals_range = {"sheetId": sheet_id, "startRowIndex": end_row - 2, "endRowIndex": end_row, "startColumnIndex": 0, "endColumnIndex": end_col}
    header_plus_data_range = {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": last_data_row_1based, "startColumnIndex": 0, "endColumnIndex": end_col}
    band_left = {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": last_data_row_1based, "startColumnIndex": 0, "endColumnIndex": (gidx if gidx is not None else end_col)}
    band_right = {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": last_data_row_1based,
                  "startColumnIndex": (gidx + 1 if gidx is not None else end_col), "endColumnIndex": end_col}
    full_table_range = {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": end_row, "startColumnIndex": 0, "endColumnIndex": end_col}

    # Column index helpers
    def col_1_based(col_name: str) -> int | None:
        try:
            return header.index(col_name) + 1
        except ValueError:
            return None

    # Number formatting applies to all numeric columns except id/name/gap
    skip_for_number = {"Member_ID", "Member_Name", GAP_COL}
    numeric_cols_1 = [i + 1 for i, c in enumerate(header) if c not in skip_for_number]

    # Conditional threshold: Day columns + AVG/d (data rows only).
    day_cols_1 = [col_1_based(c) for c in dcols]
    day_cols_1 = [c1 for c1 in day_cols_1 if c1 is not None]

    def col_range_rows(start_row_1, end_row_1, col_1):
        return {"sheetId": sheet_id, "startRowIndex": start_row_1 - 1, "endRowIndex": end_row_1,
                "startColumnIndex": col_1 - 1, "endColumnIndex": col_1}

    numeric_ranges_all = [col_range_rows(2, end_row, c1) for c1 in numeric_cols_1]
    numeric_ranges_data_days = [col_range_rows(2, last_data_row_1based, c1) for c1 in day_cols_1]

    # NEW: add AVG/d to the threshold-based red rule (data rows only)
    avgd_col_1 = col_1_based("AVG/d")
    numeric_ranges_data = list(numeric_ranges_data_days)
    if avgd_col_1 is not None:
        numeric_ranges_data.append(col_range_rows(2, last_data_row_1based, avgd_col_1))

    blue_fill  = {"red": 0.31, "green": 0.51, "blue": 0.74}
    white_font = {"red": 1, "green": 1, "blue": 1}
    red_fill   = {"red": 1.00, "green": 0.78, "blue": 0.81}
    grey_fill  = {"red": 0.75, "green": 0.75, "blue": 0.75}
    band_light = {"red": 0.86, "green": 0.92, "blue": 0.97}
    band_very  = {"red": 0.95, "green": 0.97, "blue": 0.98}
    number_format = {"type": "NUMBER", "pattern": "#,##0"}

    requests = [
        {"setBasicFilter": {"filter": {"range": header_plus_data_range}}},

        # Header styling
        {
            "repeatCell": {
                "range": header_range,
                "cell": {"userEnteredFormat": {
                    "backgroundColor": blue_fill,
                    "textFormat": {"bold": True, "foregroundColor": white_font},
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE"
                }},
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
            }
        },

        # Style both "Total" and "Day AVG" rows
        {
            "repeatCell": {
                "range": totals_range,
                "cell": {"userEnteredFormat": {"backgroundColor": blue_fill, "textFormat": {"bold": True, "foregroundColor": white_font}}},
                "fields": "userEnteredFormat(backgroundColor,textFormat)"
            }
        },

        # GAP column blue & narrow
        *([
            {
                "repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": end_row, "startColumnIndex": gidx, "endColumnIndex": gidx + 1},
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
        ] if gidx is not None else []),

        # Alternating banded rows (data only)
        *([
            {"addBanding": {"bandedRange": {"range": band_left,  "rowProperties": {"firstBandColor": band_light, "secondBandColor": band_very}}}}
        ] if gidx is None or gidx > 0 else []),
        *([
            {"addBanding": {"bandedRange": {"range": band_right, "rowProperties": {"firstBandColor": band_light, "secondBandColor": band_very}}}}
        ] if gidx is not None and gidx + 1 < end_col else []),

        # Number formatting for all numeric columns (AVG/d, Day N, Total)
        *[
            {"repeatCell": {"range": r, "cell": {"userEnteredFormat": {"numberFormat": number_format}}, "fields": "userEnteredFormat.numberFormat"}}
            for r in numeric_ranges_all
        ],

        # Conditional red (below threshold) — Day N columns + AVG/d, data rows only
        *([{
            "addConditionalFormatRule": {
                "rule": {"ranges": numeric_ranges_data,
                         "booleanRule": {"condition": {"type": "NUMBER_LESS",
                                                       "values": [{"userEnteredValue": str(threshold)}]},
                                         "format": {"backgroundColor": red_fill}}},
                "index": 0
            }
        }] if numeric_ranges_data else []),

        # Conditional grey (blanks) — ONLY for Day N columns, data rows
        *([{
            "addConditionalFormatRule": {
                "rule": {"ranges": numeric_ranges_data_days,
                         "booleanRule": {"condition": {"type": "BLANK"},
                                         "format": {"backgroundColor": grey_fill}}},
                "index": 0
            }
        }] if numeric_ranges_data_days else []),

        # Borders on all cells
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

    # Wider Member_Name (for filter icon space)
    if "Member_Name" in header:
        name_col_index = header.index("Member_Name")
        requests.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": name_col_index, "endIndex": name_col_index + 1},
                "properties": {"pixelSize": 140},
                "fields": "pixelSize"
            }
        })

    # Freeze header
    requests.append({
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
            "fields": "gridProperties.frozenRowCount"
        }
    })

    ws.spreadsheet.batch_update({"requests": requests})


# === Main ===
async def main():
    choice = pick_club()

    if choice == "ALL":
        print("\n⚡ Exporting ALL clubs: Concurrent data fetch, Sequential sheet export...\n")

        # 1. Start concurrent fetch tasks
        fetch_tasks = {
            key: asyncio.create_task(fetch_json(cfg["URL"])) 
            for key, cfg in CLUBS.items()
        }
        
        # Wait for all data fetching to complete, capturing exceptions
        await asyncio.gather(*fetch_tasks.values(), return_exceptions=True)
        
        # Define retry parameters for Google Sheets export
        MAX_EXPORT_RETRIES = 5  
        EXPORT_RETRY_DELAY = 15 
        
        # 2. Sequentially process and export data with retry logic
        for key, cfg in CLUBS.items():
            title = cfg["title"]
            
            # --- START: Export Retry Logic ---
            
            # 1. Check for data fetching errors first
            if fetch_tasks[key].exception():
                print(f"❌ {title} failed: Data fetch error: {fetch_tasks[key].exception()}")
                continue
                
            # 2. Process data before attempting export
            try:
                data = fetch_tasks[key].result() 
                df = build_dataframe(data)
            except Exception as e:
                print(f"❌ {title} failed: Data processing error: {e}")
                continue
            
            # 3. Retry loop for Sheets API (gspread)
            export_success = False
            for attempt in range(MAX_EXPORT_RETRIES):
                try:
                    print(f"→ Processing and exporting {title} (Attempt {attempt + 1}/{MAX_EXPORT_RETRIES})...")
                    export_to_gsheets(df, spreadsheet_id=SHEET_ID, sheet_title=title, threshold=cfg["THRESHOLD"])
                    
                    print(f"✅ {title} exported.")
                    export_success = True
                    break  # Success, exit the retry loop
                    
                except gspread.exceptions.APIError as e:
                    # Check for "Internal error" (500)
                    if "Internal error" in str(e) and attempt < MAX_EXPORT_RETRIES - 1:
                        print(f"⚠️ Google Sheets API Error for {title} (Attempt {attempt + 1}/{MAX_EXPORT_RETRIES}): {e}. Retrying in {EXPORT_RETRY_DELAY}s...")
                        # Use blocking time.sleep() as we are outside the primary async context loop here
                        time.sleep(EXPORT_RETRY_DELAY)
                        continue
                    else:
                        print(f"❌ {title} failed after {attempt + 1} API attempts: {e}")
                        break # Final failure or a non-retryable error, move to the next club
                
                except Exception as e:
                    # Catch any other unexpected error during export
                    print(f"❌ {title} failed during export: {e}")
                    break
            
            # --- END: Export Retry Logic ---

        print("\n🎉 All clubs export complete (check for errors above)! (Or successful)")
    
    else:
        # Single club export path (fetch_json has its own retries)
        cfg = choice
        print(f"\nSelected: {cfg['title']}\nURL: {cfg['URL']}\nSheet: {SHEET_ID}\nThreshold: {cfg['THRESHOLD']}\n")
        
        # Define retry parameters for Google Sheets export
        MAX_EXPORT_RETRIES = 5
        EXPORT_RETRY_DELAY = 15
        
        try:
            # 1. Fetch data
            data = await fetch_json(cfg["URL"])
            df = build_dataframe(data)
            
            # 2. Retry loop for export
            for attempt in range(MAX_EXPORT_RETRIES):
                try:
                    print(f"→ Processing and exporting {cfg['title']} (Attempt {attempt + 1}/{MAX_EXPORT_RETRIES})...")
                    export_to_gsheets(df, spreadsheet_id=SHEET_ID, sheet_title=cfg["title"], threshold=cfg["THRESHOLD"])
                    print(f"✅ Exported single club '{cfg['title']}' successfully!")
                    break
                except gspread.exceptions.APIError as e:
                    if "Internal error" in str(e) and attempt < MAX_EXPORT_RETRIES - 1:
                        print(f"⚠️ Google Sheets API Error for {cfg['title']} (Attempt {attempt + 1}/{MAX_EXPORT_RETRIES}): {e}. Retrying in {EXPORT_RETRY_DELAY}s...")
                        time.sleep(EXPORT_RETRY_DELAY)
                        continue
                    else:
                        print(f"❌ Exported single club '{cfg['title']}' failed after {attempt + 1} API attempts: {e}")
                        break
            
        except Exception as e:
            print(f"❌ Selected club '{cfg['title']}' failed: {e}")


if __name__ == "__main__":
    asyncio.run(main())
    input("Press Enter to close the terminal...") 