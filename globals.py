from dotenv import load_dotenv
import os

# Load .env if exists
load_dotenv()

# Defaults (safe to commit)
EXCEL_NAME = os.getenv("EXCEL_NAME", "chronogenesis_endgame_export.xlsx")
URL = os.getenv("URL", "https://chronogenesis.net/club_profile?circle_id=endgame")
EDGE_PATH = os.getenv("EDGE_PATH", r"C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe")
THRESHOLD = int(os.getenv("THRESHOLD", "1800000"))
