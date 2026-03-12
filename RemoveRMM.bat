@echo off
:: ==========================
:: RMM Removal One-Click
:: ==========================

set SCRIPTURL=https://raw.githubusercontent.com/infostreamPB/RMM/main/RemoveRMM.ps1
set SCRIPTFILE=%TEMP%\RemoveRMM.ps1

echo Downloading RMM removal script...

:: One-line PowerShell download and TLS 1.2 forcing
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%SCRIPTURL%' -OutFile '%SCRIPTFILE%' -UseBasicParsing"

:: Check if download succeeded
if exist "%SCRIPTFILE%" (
    echo Running RMM removal script...
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPTFILE%"
    echo Script finished.
    del "%SCRIPTFILE%" >nul 2>&1
) else (
    echo Failed to download the script. Check network or TLS settings.
)

pause
