@echo off
set "CHROME_EXE=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME_EXE%" set "CHROME_EXE=C:\Program Files\Google\Chrome\Application\chrome.exe"
set "SM_PROFILE=%LOCALAPPDATA%\RoadersSiteMinderChrome"
echo Starting Roaders Plus SiteMinder Dedicated Chrome (Port 9223)...
start "" "%CHROME_EXE%" --remote-debugging-port=9223 --no-first-run --no-default-browser-check --user-data-dir="%SM_PROFILE%" "https://app.siteminder.com/web/cm/properties/1bd893db-536e-44df-80de-7395edace58e/dashboard"
