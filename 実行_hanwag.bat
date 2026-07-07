@echo off
chcp 932 > nul
cd /d "%~dp0"

if not defined NO_GUI set NO_GUI=1

echo ==========================================
echo  HANWAG
echo ==========================================
python run_automation_hanwag.py
if %ERRORLEVEL% neq 0 (
    echo [ERROR] HANWAG automation failed. Stopping.
    goto :error_end
)

echo.
echo --- verify ---
python verify_all_hanwag.py
if %ERRORLEVEL% neq 0 (
    echo [ERROR] HANWAG verification failed. Stopping.
    goto :error_end
)
echo.

:success_end
echo ==========================================
echo  Done. (Process completed successfully)
echo ==========================================
pause
exit /b 0

:error_end
echo ==========================================
echo  Error or mismatch occurred. Please check the logs.
echo ==========================================
pause
exit /b 1
