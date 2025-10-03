import asyncio, json, os, sys
from pathlib import Path
import pandas as pd
import zendriver as zd

from globals import THRESHOLD, EXCEL_NAME, URL

async def main():
    # --- Grab JSON ---
    browser = await zd.start(
        browser="edge",
        browser_executable_path="C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    )
    page = await browser.get("https://google.com")
    async with page.expect_request(r".*\/api\/.*") as req:
        await page.get(URL)   
        await req.value
        body, _ = await req.response_body
    await browser.stop()

    text = body.decode("utf-8", errors="replace") if isinstance(body, (bytes, bytearray)) else str(body)
    data = json.loads(text)

    # --- Build & pivot ---
    df = pd.json_normalize(data.get("club_friend_history") or [])
    for c in ["friend_viewer_id", "friend_name", "actual_date", "adjusted_fan_gain_cumulative"]:
        if c not in df.columns: 
            df[c] = pd.NA

    df = (
        df.assign(day_col=lambda d: "Day " + d["actual_date"].astype(str))
          .pivot_table(index=["friend_viewer_id","friend_name"],
                       columns="day_col",
                       values="adjusted_fan_gain_cumulative",
                       aggfunc="first")
          .reset_index()
    )
    df.columns.name = None
    day_cols = sorted(
        [c for c in df.columns if isinstance(c, str) and c.startswith("Day ")],
        key=lambda x: int(x.split()[1]) if x.split()[1].isdigit() else 0
    )

    # Add optional Data
    df["AVG/d"] = df[day_cols].mean(axis=1).round(0)
    df = df[["friend_viewer_id","friend_name","AVG/d"] + day_cols]

    # Rename columns for clarity
    df = df.rename(columns={
        "friend_viewer_id": "Member_ID",
        "friend_name": "Member_Name"
    })

    # Keep IDs and names as string, numbers stay numeric
    text_cols = ["Member_ID", "Member_Name"]
    for col in df.columns:
        if col in text_cols:
            df[col] = df[col].fillna("").astype(str)
        else:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # --- Export to Excel ---
    if getattr(sys, 'frozen', False):  # running as .exe
        base_path = Path(sys.executable).parent
    else:  # running as .py
        base_path = Path(__file__).parent

    excel_path = str((base_path / EXCEL_NAME).resolve())
    with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
        sheet = "data"
        df.to_excel(writer, sheet_name=sheet, index=False)

        ws = writer.sheets[sheet]
        nrows, ncols = df.shape

        # Define border format
        border_fmt = writer.book.add_format({"border": 1})

        # Apply border
        ws.conditional_format(0, 0, nrows, ncols-1, {
            "type": "formula",
            "criteria": "TRUE",
            "format": border_fmt
        })

        # Conditional formatting: red if below THRESHOLD
        ws.conditional_format(1, 2, df.shape[0], ncols-1, {
            "type": "cell",
            "criteria": "<",
            "value": THRESHOLD,
            "format": writer.book.add_format({"font_color": "red"})
        })
        
        # Set column widths
        ws.set_column(0, 0, 20)  # Member_ID
        ws.set_column(1, 1, 18)  # Member_name
        if ncols > 2: 
            ws.set_column(2, ncols-1, 12)  

        ws.freeze_panes(1, 0)

    os.startfile(excel_path)

if __name__ == "__main__":
    asyncio.run(main())
