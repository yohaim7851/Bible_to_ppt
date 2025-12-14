@echo off

REM Change directory to the script's location
cd /d "%~dp0"

echo Starting python script...

python generate_ppt.py

echo Script finished. Press any key to exit.
pause
