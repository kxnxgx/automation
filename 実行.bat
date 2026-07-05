@echo off
cd /d "%~dp0"
python run_automation.py 2>&1
if %ERRORLEVEL% equ 0 (
    echo.
    echo -------------------------------------------------------
    echo  自動数値検証 (CSVの合計値と一致しているかチェック)
    echo -------------------------------------------------------
    python verify_all.py
)
echo.
echo ==========================================
echo Exit code: %ERRORLEVEL%
echo ==========================================
pause