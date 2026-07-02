@echo off
cd /d "%%~dp0"
echo Starting Invoice Generator...
start http://localhost:5000
python app.py
pause
