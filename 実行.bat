@echo off
cd /d "%~dp0"
"C:\Users\kesuzuki\AppData\Local\Python\pythoncore-3.14-64\python.exe" run_automation.py 2>&1
echo.
echo ==========================================
echo Exit code: %ERRORLEVEL%
echo ==========================================
pause