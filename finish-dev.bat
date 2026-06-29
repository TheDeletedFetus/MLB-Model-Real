@echo off
cd /d "C:\Users\rwest\OneDrive\Desktop\MLB Model"

echo =====================================
echo   MLB MODEL - FINISH DEVELOPMENT
echo =====================================
echo.

echo Pulling latest Apps Script code...
clasp pull

echo.
echo Staging Git changes...
git add -A

git diff --cached --quiet
if %errorlevel%==0 (
    echo.
    echo No changes detected.
    pause
    exit /b
)

for /f "tokens=2 delims==" %%A in ('findstr /C:"PROJECT_VERSION =" Version.js') do set VERSION_RAW=%%A

set VERSION=%VERSION_RAW:"=%
set VERSION=%VERSION:;=%
set VERSION=%VERSION: =%

echo.
echo Creating commit...
git commit -m "%VERSION%"

echo.
echo Pushing to GitHub...
git push

echo.
echo =====================================
echo GitHub Backup Complete
echo =====================================
pause