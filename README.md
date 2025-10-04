# **📄 Uma Club Tracking**

Chronogenesis Data Exporter

This project fetches club friend history data from https://chronogenesis.net/ and exports it into a clean, formatted Excel file with borders, average calculation, and conditional formatting.

**Preview**
![preview](assets/preview.png)

For the **Endless** community, simply use the `.exe` file available in **[Releases](../../releases)** — no setup required.

**⚙️ Setup**

1. Click the green Code button (top-right) → Download ZIP.
2. Extract the folder somewhere on your computer.
3. Open globals.py and update any values you want, for example:

```
   URL = "https://chronogenesis.net/club_profile?circle_id=endgame"
   THRESHOLD = 1800000
   EXCEL_NAME = "chronogenesis_endgame_export.xlsx"
```

**▶️ Usage**  
Just double-click `Script_run.bat` from File Explorer.

**🛠 Build to EXE (Windows only)**  
To package into a single .exe:
`python -m PyInstaller --onefile main.py`  
The executable will be in dist/main.exe.

![hehe](assets/image.png)
