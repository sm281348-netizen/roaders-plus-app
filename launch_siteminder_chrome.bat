@echo off
chcp 65001 >nul
echo ====================================================================
echo   啟動「站前館 SiteMinder 05:00 全自動同步專用」獨立 Chrome 視窗
echo   說明：此視窗使用獨立設定檔與 Port 9223，不影響您的日常工作分頁，
echo         只要維持 SiteMinder 登入狀態，每日 05:00 排程即可免彈窗唯讀同步。
echo ====================================================================
set "CHROME_EXE=C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME_EXE%" set "CHROME_EXE=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
set "SM_PROFILE=%LOCALAPPDATA%\RoadersSiteMinderChrome"
start "" "%CHROME_EXE%" --remote-debugging-port=9223 --user-data-dir="%SM_PROFILE%" "https://app.siteminder.com/web/cm/properties/1bd893db-536e-44df-80de-7395edace58e/dashboard"
