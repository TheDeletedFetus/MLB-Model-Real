@echo off
cd /d "C:\Users\rwest\OneDrive\Desktop\MLB Model"

echo =====================================
echo   MLB MODEL - FINISH DEVELOPMENT
echo =====================================
echo.

git add -A

git diff --cached --quiet
if %errorlevel%==0 (
    echo No changes detected.
    pause
    exit /b
)

for /f "tokens=2 delims==" %%A in ('findstr /C:"PROJECT_VERSION =" Version.js') do set VERSION_RAW=%%A

set VERSION=%VERSION_RAW:"=%
set VERSION=%VERSION:;=%
set VERSION=%VERSION: =%

echo.
echo Committing...
git commit -m "%VERSION%"

echo.
echo Pushing to GitHub...
git push

echo.
echo Pushing to Google Apps Script...
clasp push

echo.
echo =====================================
echo Sync Complete
echo =====================================
pause
