@echo off
chcp 65001 >nul
cd /d "C:\Users\88691\Desktop\roaders-plus-app"
echo ====================================================================
echo   路徒 PLUS 站前館 - SiteMinder 100%% 唯讀快照一鍵同步並推播雲端
echo   (若 Chrome 跳出「允許遠端偵錯」請點選允許，全程 100%% 唯讀零修改)
echo ====================================================================
"C:\Users\88691\AppData\Local\Python\pythoncore-3.14-64\python.exe" siteminder_cdp_sync.py --interactive --windows 3 --git-push
echo.
pause
