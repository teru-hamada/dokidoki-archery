@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Github\dokidoki-archery\shell\score_receiver.ps1" -Port 5000 -LowScoreThreshold 15
endlocal
