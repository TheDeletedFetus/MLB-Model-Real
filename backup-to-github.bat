@echo off
cd /d "C:\Users\rwest\OneDrive\Desktop\MLB Model"

for /f "tokens=2 delims==" %%A in ('findstr /C:"MODEL_VERSION =" Version.js') do set VERSION_RAW=%%A

set VERSION=%VERSION_RAW:"=%
set VERSION=%VERSION:;=%
set VERSION=%VERSION: =%

git add -A

git diff --cached --quiet
if %errorlevel%==0 (
    echo No changes to commit.
    pause
    exit /b
)

git commit -m "%VERSION% daily backup - %date% %time%"
git push

pause