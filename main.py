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
    """Print club list and return chosen config (defaults to '1' if invalid)."""
    print("=== Choose a club to export ===")
    for key, cfg in CLUBS.items():
        print(f"{key}. {cfg['title']}")
    choice = input("Enter 1–7: ").strip()
    if choice not in CLUBS:
        print("Invalid choice, defaulting to 1.")
        choice = "1"
    return CLUBS[choice]


# === Paths and runtime ===
def resolve_base_dir() -> Path:
    """Directory for the executable or the script file."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


# === Data fetch ===
async def fetch_json(URL: str):
    """Open a temporary browser, navigate, and capture the first /api/ response as JSON."""
    browser = await zd.start(
        browser="edge",  # change if using a different browser
        browser_executable_path="C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"  # adjust path if needed
    )
    page = await browser.get("https://google.com")

    # Wait for the first matching API request triggered by navigating to URL
    async with page.expect_request(r".*/api/.*") as req:
        await page.get(URL)
        await req.value
        body, _ = await req.response_body

    await browser.stop()
    text = body.decode("utf-8", errors="replace") if isinstance(body, (bytes, bytearray)) else str(body)
    return json.loads(text)


# === DataFrame processing ===
def build_dataframe(data: dict) -> tuple[pd.DataFrame, list[str]]:
    """Flatten JSON, pivot daily values to wide columns, and compute AVG/d."""
    df = pd.json_normalize(data.get("club_friend_history") or [])
    required = ["friend_viewer_id", "friend_name", "actual_date", "adjusted_interpolated_fan_gain"]
    for c in required:
        if c not in df.columns:
            df[c] = pd.NA

    # Wide table: one row per member, one column per day
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

    # Sort day columns numerically (Day 1, Day 2, ...)
    def _day_key(x: str):
        part = x.split(maxsplit=1)[1] if " " in x else x
        try:
            return int(part)
        except ValueError:
            return part

    day_cols = sorted([c for c in df.columns if str(c).startswith("Day ")], key=_day_key)

    # Average per member
    df["AVG/d"] = df[day_cols].mean(axis=1).round(0) if day_cols else 0
    df = df[["friend_viewer_id", "friend_name", "AVG/d"] + day_cols]
    df = df.rename(columns={"friend_viewer_id": "Member_ID", "friend_name": "Member_Name"})

    # Types: keep IDs/Names as string; numeric for metrics
    for col in df.columns:
        if col in ["Member_ID", "Member_Name"]:
            df[col] = df[col].fillna("").astype(str)
        else:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df, day_cols


# === Excel export ===
def export_excel(df: pd.DataFrame, excel_path: str, threshold: int, sheet_name: str):
    """Write DataFrame to Excel with widths, totals, and conditional formats."""
    def col_idx_safe(columns, name):
        try:
            return columns.get_loc(name)
        except KeyError:
            return None

    def day_columns(columns):
        return [c for c in columns if isinstance(c, str) and c.startswith("Day ")]

    GAP_COL = " "
    dcols = day_columns(df.columns)

    # Insert Total and a visual gap column
    if dcols:
        df["Total"] = df[dcols].apply(pd.to_numeric, errors="coerce").sum(axis=1, min_count=1)
        df.insert(df.columns.get_loc("Total"), GAP_COL, "")

    # Append bottom totals row
    totals = {}
    for c in df.columns:
        if c == "Member_Name":
            totals[c] = "Total"
        elif c in ("Member_ID", GAP_COL):
            totals[c] = ""
        else:
            totals[c] = pd.to_numeric(df[c], errors="coerce").sum(min_count=1)
    df = pd.concat([df, pd.DataFrame([totals])], ignore_index=True)

    # Write Excel and format
    with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]

        nrows, ncols = df.shape
        last_row = nrows
        last_data_row = nrows - 1

        # Formats
        book = writer.book
        fmt_border_all = book.add_format({"border": 1})
        fmt_red_fill = book.add_format({"bg_color": "#FFC7CE", "border": 1, "border_color": "red"})
        fmt_text = book.add_format({"num_format": "@"})
        fmt_bold_row = book.add_format({"bold": True, "border": 1})
        fmt_bold_cell = book.add_format({"bold": True, "border": 1})
        fmt_num = book.add_format({"num_format": "#,##0"})
        fmt_gap_blue = book.add_format({"bg_color": "#D9E1F2"})
        fmt_blank_grey = book.add_format({"bg_color": "#F2F2F2"})
        fmt_header = book.add_format({"bold": True, "bg_color": "#C6EFCE", "border": 1})

        # Column widths & base formats
        ws.set_column(0, 0, 20, fmt_text)
        ws.set_column(1, 1, 18, fmt_text)
        ws.set_column(2, ncols - 1, 12)

        gidx = col_idx_safe(df.columns, GAP_COL)
        if gidx is not None:
            ws.set_column(gidx, gidx, 2)

        # Header row (re-written to apply header format)
        for c in range(ncols):
            ws.write(0, c, df.columns[c], fmt_gap_blue if c == gidx else fmt_header)

        # Add filter over used range (exclude bottom Total row)
        if nrows > 1:
            ws.autofilter(0, 0, last_data_row - 1, ncols - 1)
            
        # Thin borders
        ws.conditional_format(1, 0, last_row, ncols - 1, {
            "type": "formula", "criteria": "TRUE", "format": fmt_border_all
        })

        # Color the gap column
        if gidx is not None:
            ws.conditional_format(1, gidx, last_row, gidx, {
                "type": "formula", "criteria": "TRUE", "format": fmt_gap_blue
            })

        # Numeric formats for day and AVG columns
        if dcols:
            first_day = df.columns.get_loc(dcols[0])
            last_day = df.columns.get_loc(dcols[-1])
            ws.conditional_format(1, first_day, last_row, last_day, {
                "type": "no_blanks", "format": fmt_num
            })

        avg_idx = col_idx_safe(df.columns, "AVG/d")
        if avg_idx is not None:
            ws.conditional_format(1, avg_idx, last_row, avg_idx, {
                "type": "no_blanks", "format": fmt_num
            })

        total_col_idx = col_idx_safe(df.columns, "Total")
        if total_col_idx is not None:
            ws.conditional_format(1, total_col_idx, last_row, total_col_idx, {
                "type": "no_blanks", "format": fmt_num
            })

        # Shade blanks
        ws.conditional_format(1, 0, last_row, ncols - 1, {
            "type": "blanks", "format": fmt_blank_grey
        })

        # Highlight below threshold
        if dcols and nrows > 1:
            first_day = df.columns.get_loc(dcols[0])
            last_day = df.columns.get_loc(dcols[-1])
            ws.conditional_format(1, first_day, last_data_row - 1, last_day, {
                "type": "cell", "criteria": "<", "value": threshold, "format": fmt_red_fill
            })
        if avg_idx is not None and nrows > 1:
            ws.conditional_format(1, avg_idx, last_data_row - 1, avg_idx, {
                "type": "cell", "criteria": "<", "value": threshold, "format": fmt_red_fill
            })

        # Bold totals
        ws.conditional_format(last_row, 0, last_row, ncols - 1, {
            "type": "formula", "criteria": "TRUE", "format": fmt_bold_row
        })
        if total_col_idx is not None:
            ws.conditional_format(1, total_col_idx, last_row, total_col_idx, {
                "type": "formula", "criteria": "TRUE", "format": fmt_bold_cell
            })

        ws.freeze_panes(1, 0)


# === Helpers ===
def open_excel_windows(excel_path: str):
    """Open the generated file on Windows."""
    os.startfile(excel_path)


# === Main ===
async def main():
    """Orchestrate: choose club → fetch → transform → export → open."""
    cfg = pick_club()
    URL = cfg["URL"]
    EXCEL_NAME = cfg["EXCEL_NAME"]
    THRESHOLD = cfg["THRESHOLD"]

    print(f"\nSelected: {cfg['title']}")
    print(f"URL: {URL}")
    print(f"Excel: {EXCEL_NAME}")
    print(f"Threshold: {THRESHOLD}\n")

    data = await fetch_json(URL)
    df, _ = build_dataframe(data)

    base_dir = resolve_base_dir()
    excel_path = str((base_dir / EXCEL_NAME).resolve())
    export_excel(df, excel_path, THRESHOLD, cfg["title"])

    try:
        open_excel_windows(excel_path)
    except Exception as e:
        print(f"Exported to: {excel_path} (could not auto-open: {e})")


if __name__ == "__main__":
    asyncio.run(main())
