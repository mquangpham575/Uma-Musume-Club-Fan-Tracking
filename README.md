# **📄 Uma Club Tracking — Chronogenesis Exporter (Preview 3.0)**

This project fetches **club friend history data** from [ChronoGenesis](https://chronogenesis.net/)  
and automatically exports it into a **formatted Google Spreadsheet** — complete with borders, totals, averages, and conditional formatting.

**Preview**
![preview](assets/preview.png)

For the **Endless** community, you can directly use the `.exe` from **[Releases](../../releases)** — no setup needed.

---

## ⚙️ Setup

1. Click the green **Code** button → **Download ZIP**
2. Extract the folder anywhere on your computer
3. Place your **Google API credentials** file as `credentials.json` in the same directory
4. Open `globals.py` and edit these values if needed:

'''SHEET_ID = "1O09PM-hYo-H05kWWqMg71GelEpfaGrePQWzdDCKOqyU"

CLUBS = {
"1": {"title": "EndGame", "URL": "https://chronogenesis.net/club_profile?circle_id=endgame", "THRESHOLD": 1800000},
...
}'''

## ▶️ Usage

Simply double-click:
'''"Script_run.bat"'''
Then choose:
'''=== Choose a club to export ===

1. EndGame
2. AnotherClub  
   ...
3. Export ALL clubs (default)
   Enter 0–7 [default=0]:'''

- Press Enter / 0: export all clubs in parallel
- Enter a number: export a single club only
- Each club will appear as a separate sheet inside your Google Spreadsheet.

## 🧾 Export Details

- Header & totals → **bold, white text on blue background**
- Alternating light rows for readability
- Automatic borders around all cells
- Conditional colors:
  - 🔴 **Red** → value below threshold
  - ⚪ **Grey** → blank cell
- `Member_Name` column auto-sized (fits filter icon)
- Adds a **Total** column & row automatically

## ⚡ Parallel Mode

Version 3.0 now runs all clubs **asynchronously in parallel**,

reducing total runtime from 8× down to roughly 1× overall loading time.

> Each club runs in its own Edge browser instance (via zendriver).
>
> Expect brief multiple Edge popups — this is normal.

---

## 🧰 Requirements

Install dependencies:

```
pip install -r requirements.txt
```
