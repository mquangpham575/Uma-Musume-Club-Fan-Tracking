import asyncio, json, os, sys
from pathlib import Path
import pandas as pd
import zendriver as zd

from globals import CLUBS

def pick_club() -> dict:
    print("=== Choose a club to export ===")
    for key, cfg in CLUBS.items():
        print(f"{key}. {cfg['title']}")
    choice = input("Enter 1-7: ").strip()
    if choice not in CLUBS:
        print("Invalid choice, defaulting to 1.")
        choice = "1"
    return CLUBS[choice]

def resolve_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent

async def fetch_json(URL: str):
    browser = await zd.start(
        browser="edge",
        browser_executable_path="C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
    )
    page = await browser.get("https://google.com")
    async with page.expect_request(r".*\/api\/.*") as req:
        await page.get(URL)
        await req.value
        body, _ = await req.response_body
    await browser.stop()

    text = body.decode("utf-8", errors="replace") if isinstance(body, (bytes, bytearray)) else str(body)
    return json.loads(text)

def build_dataframe(data: dict) -> tuple[pd.DataFrame, list[str]]:
    # Flatten and ensure columns exist
    df = pd.json_normalize(data.get("club_friend_history") or [])
    needed = ["friend_viewer_id", "friend_name", "actual_date", "adjusted_interpolated_fan_gain"]
    for c in needed:
        if c not in df.columns:
            df[c] = pd.NA

    # Build "Day x" columns from the *daily* (interpolated) gain
    df = (
        df.assign(day_col=lambda d: "Day " + d["actual_date"].astype(str))
          .pivot_table(
              index=["friend_viewer_id", "friend_name"],
              columns="day_col",
              values="adjusted_interpolated_fan_gain",
              aggfunc="first"  # if there can be multiple rows per day, switch to "sum"
          )
          .reset_index()
    )
    df.columns.name = None

    # Sort day columns numerically if possible
    day_cols = [c for c in df.columns if isinstance(c, str) and c.startswith("Day ")]
    def _day_key(x: str):
        part = x.split(maxsplit=1)[1] if " " in x else x
        try:
            return int(part)
        except Exception:
            return part
    day_cols = sorted(day_cols, key=_day_key)

    # AVG per day (from daily values)
    df["AVG/d"] = df[day_cols].mean(axis=1).round(0) if day_cols else 0
    df = df[["friend_viewer_id", "friend_name", "AVG/d"] + day_cols]

    # Rename and set dtypes
    df = df.rename(columns={"friend_viewer_id": "Member_ID", "friend_name": "Member_Name"})
    for col in df.columns:
        if col in ["Member_ID", "Member_Name"]:
            df[col] = df[col].fillna("").astype(str)
        else:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df, day_cols

def export_excel(df: pd.DataFrame, excel_path: str, threshold: int, sheet_name: str):
    with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        nrows, ncols = df.shape

        # Formats
        border_fmt = writer.book.add_format({"border": 1})
        red_fmt = writer.book.add_format({"font_color": "red", "border": 1})
        text_fmt = writer.book.add_format({"num_format": "@"})

        # Column widths
        ws.set_column(0, 0, 20, text_fmt)  # Member_ID
        ws.set_column(1, 1, 18, text_fmt)  # Member_Name
        if ncols > 2:
            ws.set_column(2, ncols-1, 12)   # numeric

        # Border EVERY cell (including blanks): formula TRUE over full range
        ws.conditional_format(0, 0, nrows, ncols-1, {
            "type": "formula",
            "criteria": "TRUE",
            "format": border_fmt
        })

        # Red if below threshold (AVG/d + days)
        if ncols > 2:
            ws.conditional_format(1, 2, nrows, ncols-1, {
                "type": "cell",
                "criteria": "<",
                "value": threshold,
                "format": red_fmt
            })

        ws.freeze_panes(1, 0)

def open_excel_windows(excel_path: str):
    os.startfile(excel_path)  # Windows only

async def main():
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
    sheet_name = cfg["title"]
    export_excel(df, excel_path, THRESHOLD, sheet_name)

    try:
        open_excel_windows(excel_path)
    except Exception as e:
        print(f"Exported to: {excel_path} (could not auto-open: {e})")

if __name__ == "__main__":
    asyncio.run(main())
