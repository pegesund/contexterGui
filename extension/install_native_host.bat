@echo off
REM Register NorskTale native messaging host for Chrome and Edge
REM Run this as Administrator after building native_bridge.exe

set MANIFEST_PATH=%~dp0com.norsktale.bridge.json

REM Chrome
reg add "HKCU\Software\Google\Chrome\NativeMessagingHosts\com.norsktale.bridge" /ve /t REG_SZ /d "%MANIFEST_PATH%" /f

REM Edge
reg add "HKCU\Software\Microsoft\Edge\NativeMessagingHosts\com.norsktale.bridge" /ve /t REG_SZ /d "%MANIFEST_PATH%" /f

echo Native messaging host registered for Chrome and Edge.
echo.
echo Next steps:
echo 1. Build: cargo build --release --bin native_bridge
echo 2. Load the extension in edge://extensions (Developer mode, Load unpacked)
echo 3. Copy the extension ID into com.norsktale.bridge.json allowed_origins
echo 4. Re-run this script
pause
