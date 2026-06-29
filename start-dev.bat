@echo off
cd /d "C:\Users\rwest\OneDrive\Desktop\MLB Model"

echo =====================================
echo   MLB MODEL - START DEVELOPMENT
echo =====================================
echo.

echo Pulling latest GitHub changes...
git pull

echo.
echo Pulling latest Apps Script changes...
clasp pull

echo.
echo Opening project...
start "" .

echo.
echo Ready to code.
pause
