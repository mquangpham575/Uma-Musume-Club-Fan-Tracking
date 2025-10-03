@echo off
echo Installing dependencies...
pip install -r requirements.txt >nul 2>&1

echo.
echo Waiting for data...
python main.py
