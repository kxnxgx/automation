@echo off
chcp 932 > nul
cd /d "%~dp0"

if not defined NO_GUI set NO_GUI=1

echo.
echo ==========================================
echo  [1/3] FRV
echo ==========================================
python run_automation.py
if %ERRORLEVEL% neq 0 (
    echo [ERROR] FRV automation failed. Stopping.
    goto :error_end
)
python verify_all.py
if %ERRORLEVEL% neq 0 (
    echo [ERROR] FRV verification failed. Stopping.
    goto :error_end
)
echo.

echo ==========================================
echo  [2/3] tennen
echo ==========================================
python run_automation_tennen.py
if %ERRORLEVEL% neq 0 (
    echo [ERROR] tennen automation failed. Stopping.
    goto :error_end
)
python verify_all_tennen.py
if %ERRORLEVEL% neq 0 (
    echo [ERROR] tennen verification failed. Stopping.
    goto :error_end
)
echo.

echo ==========================================
echo  [3/3] HANWAG
echo ==========================================
python run_automation_hanwag.py
if %ERRORLEVEL% neq 0 (
    echo [ERROR] HANWAG automation failed. Stopping.
    goto :error_end
)
python verify_all_hanwag.py
if %ERRORLEVEL% neq 0 (
    echo [ERROR] HANWAG verification failed. Stopping.
    goto :error_end
)
echo.

:success_end
echo ==========================================
echo  Done. (All process completed successfully)
echo ==========================================
pause
exit /b 0

:error_end
echo ==========================================
echo  Error or mismatch occurred. Please check the logs.
echo ==========================================
pause
exit /b 1