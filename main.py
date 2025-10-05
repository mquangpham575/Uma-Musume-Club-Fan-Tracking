import asyncio
import json
import os
import sys
from pathlib import Path
import pandas as pd
import zendriver as zd

from globals import CLUBS  # external config for club list/URLs/thresholds


# === Club selection ===
def pick_club() -> dict:
    print("=== Choose a club to export ===")
    for key, cfg in CLUBS.items():
        print(f"{key}. {cfg['title']}")
    choice = input("Enter 1–7: ").strip()
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

    # Sort Day N columns numerically
    def _day_key(x: str):
        if not isinstance(x, str) or not x.startswith("Day "): return x
        try:
            return int(x.split(maxsplit=1)[1])
        except Exception:
            return x
    day_cols = sorted([c for c in df.columns if isinstance(c, str) and c.startswith("Day ")], key=_day_key)

    # AVG/d, rename ID/Name, cast numeric
    df["AVG/d"] = df[day_cols].mean(axis=1).round(0) if day_cols else 0
    df = df[["friend_viewer_id", "friend_name", "AVG/d"] + day_cols].rename(
        columns={"friend_viewer_id": "Member_ID", "friend_name": "Member_Name"}
    )
    df["Member_ID"] = df["Member_ID"].fillna("").astype(str)
    df["Member_Name"] = df["Member_Name"].fillna("").astype(str)
    for c in df.columns:
        if c not in ("Member_ID", "Member_Name"):
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Sort rows: largest AVG/d first; tie-break by name (stable sort)
    df = df.sort_values(["AVG/d", "Member_Name"], ascending=[False, True], kind="mergesort").reset_index(drop=True)
    return df

# === Excel export ===
def export_excel(df: pd.DataFrame, excel_path: str, threshold: int, sheet_name: str):
    import re

    def col_idx(columns, name):
        try:
            return columns.get_loc(name)
        except KeyError:
            return None

    def day_columns(columns):
        return [c for c in columns if isinstance(c, str) and c.startswith("Day ")]

    GAP_COL = " "
    dcols = day_columns(df.columns)

    # Add per-row Total and insert gap column before it
    if dcols:
        df = df.copy()
        df["Total"] = df[dcols].sum(axis=1, min_count=1)
        df.insert(df.columns.get_loc("Total"), GAP_COL, "")

    # Precompute bottom totals (label under Member_Name; blank under id/gap)
    bottom_totals = {}
    for c in df.columns:
        if c == "Member_Name":
            bottom_totals[c] = "Total"
        elif c in ("Member_ID", GAP_COL):
            bottom_totals[c] = ""
        else:
            bottom_totals[c] = pd.to_numeric(df[c], errors="coerce").sum(min_count=1)

    with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        book = writer.book

        # Layout numbers
        nrows, ncols = df.shape
        first_data_row = 1
        last_data_row = first_data_row + nrows - 1
        totals_row = last_data_row + 1

        # Colors/styles
        header_color = "#4F81BD"   # Excel Table Style Medium 2 header blue
        header_font  = "#FFFFFF"

        fmt_num        = book.add_format({"num_format": "#,##0"})
        fmt_text       = book.add_format({"num_format": "@"})
        fmt_bold_cell  = book.add_format({"bold": True, "border": 1})
        fmt_border_all = book.add_format({"border": 1})
        fmt_blank_grey = book.add_format({"bg_color": "#BFBFBF"})
        fmt_red_fill   = book.add_format({"bg_color": "#FFC7CE", "border": 1, "border_color": "red"})
        fmt_gap_column = book.add_format({"bg_color": header_color})
        fmt_total_row  = book.add_format({
            "bg_color": header_color, "font_color": header_font, "bold": True,
            "border": 1, "align": "center", "valign": "vcenter"
        })

        # Column widths
        ws.set_column(0, 0, 20, fmt_text)        # Member_ID
        ws.set_column(1, 1, 18, fmt_text)        # Member_Name
        ws.set_column(2, ncols - 1, 12)

        gidx = col_idx(df.columns, GAP_COL)
        if gidx is not None:
            ws.set_column(gidx, gidx, 2)

        # Create table(s), excluding the gap column so it won’t inherit banding
        table_name = re.sub(r"[^A-Za-z0-9_]", "_", f"tbl_{sheet_name or 'Sheet'}") or "tbl_Data"
        if gidx is not None:
            ws.add_table(0, 0, last_data_row, gidx - 1, {
                "name": table_name + "_L", "style": "Table Style Medium 2",
                "columns": [{"header": str(h)} for h in df.columns[:gidx]],
                "autofilter": True, "banded_rows": True
            })
            if gidx < ncols - 1:
                ws.add_table(0, gidx + 1, last_data_row, ncols - 1, {
                    "name": table_name + "_R", "style": "Table Style Medium 2",
                    "columns": [{"header": str(h)} for h in df.columns[gidx + 1:]],
                    "autofilter": True, "banded_rows": True
                })
        else:
            ws.add_table(0, 0, last_data_row, ncols - 1, {
                "name": table_name, "style": "Table Style Medium 2",
                "columns": [{"header": str(h)} for h in df.columns],
                "autofilter": True, "banded_rows": True
            })

        # Borders and blanks (data only)
        ws.conditional_format(first_data_row, 0, last_data_row, ncols - 1,
                              {"type": "formula", "criteria": "TRUE", "format": fmt_border_all})

        # Gap column fill for DATA rows only (leave header, totals untouched)
        if gidx is not None:
            ws.conditional_format(0, gidx, last_data_row, gidx,
                                  {"type": "formula", "criteria": "TRUE", "format": fmt_gap_column, "stop_if_true": True})

            if gidx > 0:
                ws.conditional_format(first_data_row, 0, last_data_row, gidx - 1,
                                      {"type": "blanks", "format": fmt_blank_grey})
            if gidx < ncols - 1:
                ws.conditional_format(first_data_row, gidx + 1, last_data_row, ncols - 1,
                                      {"type": "blanks", "format": fmt_blank_grey})
        else:
            ws.conditional_format(first_data_row, 0, last_data_row, ncols - 1,
                                  {"type": "blanks", "format": fmt_blank_grey})

        # Numeric formats + threshold highlights (data only)
        if dcols:
            c0, c1 = df.columns.get_loc(dcols[0]), df.columns.get_loc(dcols[-1])
            ws.conditional_format(first_data_row, c0, last_data_row, c1,
                                  {"type": "no_blanks", "format": fmt_num})
            ws.conditional_format(first_data_row, c0, last_data_row, c1,
                                  {"type": "cell", "criteria": "<", "value": threshold, "format": fmt_red_fill})

        avg_idx = col_idx(df.columns, "AVG/d")
        if avg_idx is not None:
            ws.conditional_format(first_data_row, avg_idx, last_data_row, avg_idx,
                                  {"type": "no_blanks", "format": fmt_num})
            ws.conditional_format(first_data_row, avg_idx, last_data_row, avg_idx,
                                  {"type": "cell", "criteria": "<", "value": threshold, "format": fmt_red_fill})

        total_col_idx = col_idx(df.columns, "Total")
        if total_col_idx is not None:
            ws.conditional_format(first_data_row, total_col_idx, last_data_row, total_col_idx,
                                  {"type": "no_blanks", "format": fmt_num})

        # Append totals row (outside table) and style it like the header
        if "Member_Name" in df.columns:
            ws.write(totals_row, df.columns.get_loc("Member_Name"), "Total", fmt_bold_cell)
        for j, col_name in enumerate(df.columns):
            if col_name in ("Member_ID", "Member_Name", GAP_COL):
                continue
            val = bottom_totals.get(col_name, "")
            if pd.isna(val):
                continue
            if isinstance(val, (int, float)):
                ws.write_number(totals_row, j, float(val), fmt_bold_cell)
            else:
                ws.write(totals_row, j, val, fmt_bold_cell)

        ws.conditional_format(totals_row, 0, totals_row, ncols - 1,
                              {"type": "formula", "criteria": "TRUE", "format": fmt_total_row})

        # Keep Total column bold through the bottom
        if total_col_idx is not None:
            ws.conditional_format(first_data_row, total_col_idx, totals_row, total_col_idx,
                                  {"type": "formula", "criteria": "TRUE", "format": fmt_bold_cell})

        ws.freeze_panes(1, 0)


# === Helpers ===
def open_excel_windows(excel_path: str):
    os.startfile(excel_path)


# === Main ===
async def main():
    cfg = pick_club()
    URL = cfg["URL"]; EXCEL_NAME = cfg["EXCEL_NAME"]; THRESHOLD = cfg["THRESHOLD"]

    print(f"\nSelected: {cfg['title']}\nURL: {URL}\nExcel: {EXCEL_NAME}\nThreshold: {THRESHOLD}\n")

    data = await fetch_json(URL)
    df = build_dataframe(data)

    base_dir = resolve_base_dir()
    excel_path = str((base_dir / EXCEL_NAME).resolve())
    export_excel(df, excel_path, THRESHOLD, cfg["title"])

    try:
        open_excel_windows(excel_path)
    except Exception as e:
        print(f"Exported to: {excel_path} (could not auto-open: {e})")


if __name__ == "__main__":
    asyncio.run(main())
