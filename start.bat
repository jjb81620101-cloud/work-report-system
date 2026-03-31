@echo off
cd /d "%~dp0"

echo Installing packages...
pip install flask flask-cors werkzeug gspread google-auth

echo.
echo Starting server...
python server.py

pause
