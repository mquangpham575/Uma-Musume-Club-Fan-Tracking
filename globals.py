from dotenv import load_dotenv
import os

# Load .env if it exists
load_dotenv()

# Path to Microsoft Edge executable (default is 32-bit Program Files)
EDGE_PATH = os.getenv("EDGE_PATH", r"C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe")

# Threshold for conditional formatting in Excel
# Any value below this will be highlighted in red
THRESHOLD = int(os.getenv("THRESHOLD", "1800000"))

# Name of the Excel file to generate
EXCEL_NAME = os.getenv("EXCEL_NAME", "chronogenesis_endgame_export.xlsx")

# API URL for fetching Chronogenesis club friend history
URL = os.getenv("URL", "https://chronogenesis.net/club_profile?circle_id=endgame")
