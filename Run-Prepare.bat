@echo off
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -File "%~dp0Scripts\Prepare-Refunds.ps1"
if errorlevel 1 echo Preparation stopped. Read the message above. Do not use uncommitted outputs.
pause
