@echo off

echo.
echo =====================================
echo MLB MODEL GITHUB SYNC
echo =====================================
echo.

cd /d "C:\Users\rwest\OneDrive\Desktop\MLB Model"

echo Current Directory:
cd

echo.
echo Pulling latest Apps Script files...
call clasp pull

if errorlevel 1 (
    echo.
    echo ERROR: CLASP PULL FAILED
    pause
    exit /b
)

echo.
echo Adding files...
git add .

echo.
echo Checking for changes...
git diff --cached --quiet

if %errorlevel%==0 (
    echo.
    echo No changes detected.
    exit /b
)

echo.
echo Creating commit...
git commit -m "Automated Apps Script sync"

echo.
echo Pushing to GitHub...
git push origin main

echo.
echo =====================================
echo SYNC COMPLETE
echo =====================================
pause