@echo off
chcp 65001 >nul
cd /d "C:\Users\88691\Desktop\roaders-plus-app"
"C:\Users\88691\AppData\Local\Python\pythoncore-3.14-64\python.exe" run_daily_sync.py --days 3 --git-push >> logs\scheduler.log 2>&1
