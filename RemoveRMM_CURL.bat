@echo off
set SCRIPTURL=https://raw.githubusercontent.com/infostreamPB/RMM/main/RemoveRMM.ps1
set SCRIPTFILE=%TEMP%\RemoveRMM.ps1

echo Downloading RMM removal script with curl...
curl -L --tlsv1.2 -o "%SCRIPTFILE%" "%SCRIPTURL%"

if exist "%SCRIPTFILE%" (
    echo Running RMM removal script...
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPTFILE%"
    del "%SCRIPTFILE%" >nul 2>&1
    echo Script finished.
) else (
    echo Failed to download script. Check network or TLS settings.
)

pause