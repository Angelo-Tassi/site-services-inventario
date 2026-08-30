@echo off
rem Da eseguire una sola volta su ogni PC che non ha ancora openpyxl.
cd /d "%~dp0"
python -m pip install --user -r requirements.txt
pause
