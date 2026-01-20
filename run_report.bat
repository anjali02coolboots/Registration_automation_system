@echo off
echo Starting Automated Report Generation...
echo =========================================

REM Change to script directory
cd /d "C:\Users\Anjali\Desktop\Work stuff\Report 4"

REM Run Python script
python app.py

REM Check if script succeeded
if %ERRORLEVEL% EQU 0 (
    echo SUCCESS: Script completed at %date% %time%
) else (
    echo ERROR: Script failed with error code %ERRORLEVEL%
)

REM Exit with the error code (important for Task Scheduler)
exit /b %ERRORLEVEL%