@echo off
setlocal

cd /d C:\Users\rwest\MLB-Model-Real

echo ================================
echo MLB Model GitHub Auto Sync
echo ================================
echo.

echo Pulling latest Apps Script files...
call clasp pull

if errorlevel 1 (
  echo.
  echo ERROR: clasp pull failed.
  pause
  exit /b 1
)

echo.
echo Checking for changes...
git status --short

echo.
echo Adding changed files...
git add .

echo.
echo Checking whether there is anything to commit...
git diff --cached --quiet

if %errorlevel%==0 (
  echo No changes to commit.
  echo.
  echo Sync complete.
  exit /b 0
)

echo.
echo Committing changes...
git commit -m "Auto-sync Apps Script updates"

if errorlevel 1 (
  echo.
  echo ERROR: git commit failed.
  pause
  exit /b 1
)

echo.
echo Pushing to GitHub...
git push origin main

if errorlevel 1 (
  echo.
  echo ERROR: git push failed.
  pause
  exit /b 1
)

echo.
echo Sync complete.
exit /b 0