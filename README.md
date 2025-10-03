**#📄 Uma Club Tracking**

Chronogenesis Data Exporter

This project fetches club friend history data from https://chronogenesis.net/ and exports it into a clean, formatted Excel file with borders, average calculation, and conditional formatting.

**⚙️ Setup**

1. Click the green Code button (top-right) → Download ZIP.
2. Extract the folder somewhere on your computer.
3. Open globals.py and update any values you want, for example:

```
   EXCEL_NAME = "chronogenesis_endgame_export.xlsx"
   URL = "https://chronogenesis.net/club_profile?circle_id=endgame"
   EDGE_PATH = r"C:/Program Files/Microsoft/Edge/Application/msedge.exe"
   THRESHOLD = 2000000
```

4. Install dependencies:
   `pip install -r requirements.txt`

**▶️ Usage**  
Run the script with:  
`python main.py`

**🛠 Build to EXE (Windows only)**  
To package into a single .exe:
`python -m PyInstaller --onefile main.py`  
The executable will be in dist/main.exe.
