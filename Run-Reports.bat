@echo off
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -File "%~dp0Scripts\Generate-Reports.ps1"
pause
