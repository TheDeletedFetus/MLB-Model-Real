@echo off
cd /d "C:\Users\rwest\OneDrive\Desktop\MLB Model"

git add -A

git commit -m "Daily backup %date% %time%"

git push

pause