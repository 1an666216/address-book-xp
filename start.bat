@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Starting Address Book on http://127.0.0.1:8000 ...
python app.py --port 8000
pause
