@echo off
set SCRIPTURL=https://raw.githubusercontent.com/infostreamPB/RMM/main/RemoveRMM.ps1
set SCRIPTFILE=%TEMP%\RemoveRMM.ps1

echo Downloading script...
curl -L -o "%SCRIPTFILE%" "%SCRIPTURL%"

if exist "%SCRIPTFILE%" (
echo Running script...
powershell -ExecutionPolicy Bypass -NoProfile -File "%SCRIPTFILE%"
) else (
echo Failed to download script.
)

pause
