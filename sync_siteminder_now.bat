@echo off
cd /d "C:\Users\88691\Desktop\roaders-plus-app"
echo Running SiteMinder Read-Only Snapshot Sync...
"C:\Users\88691\AppData\Local\Python\pythoncore-3.14-64\python.exe" siteminder_cdp_sync.py --interactive --windows 3 --git-push
echo.
pause
