@echo off
echo Starting CM360 Audit Testing...
echo.

cd /d "c:\Users\EvSchneider\cm360-audit-default"

echo Pushing code to Google Apps Script...
powershell -ExecutionPolicy Bypass -Command "clasp push"
echo.

echo Running date parsing test...
powershell -ExecutionPolicy Bypass -Command "clasp run testDateParsing"
echo.

echo Running full audit...
powershell -ExecutionPolicy Bypass -Command "clasp run runAudit"
echo.

echo Checking logs...
powershell -ExecutionPolicy Bypass -Command "clasp logs"
echo.

echo Test complete!
pause
