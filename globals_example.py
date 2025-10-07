# === Google Sheets target ===
# The spreadsheet where all club data will be exported
SHEET_ID = "1O09PM-hYo-H05kWWqMg71GelEpfaGrePQWzdDCKOqyU"

# === Club list configuration ===
# Each club defines its own Chronogenesis profile URL and threshold (minimum daily average)
CLUBS = {
    "1": {
        "title": "EndGame",
        "URL": "https://chronogenesis.net/club_profile?circle_id=endgame",
        "THRESHOLD": 1800000
    },
    "2": {
        "title": "EverNight",
        "URL": "https://chronogenesis.net/club_profile?circle_id=evernight",
        "THRESHOLD": 1500000
    },
    "3": {
        "title": "Chrono",
        "URL": "https://chronogenesis.net/club_profile?circle_id=chrono",
        "THRESHOLD": 1200000
    },
    # Add more clubs as needed...
}
