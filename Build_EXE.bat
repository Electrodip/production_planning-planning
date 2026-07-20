@echo off
cd /d "%~dp0"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --onefile --windowed --name ElectroDipProductionPlanner app.py
echo.
echo EXE created in the dist folder.
pause
