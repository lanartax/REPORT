@echo off
cd /d "C:\Users\Inventaris SJS\Documents\REPORT"
tasklist /fi "imagename eq node.exe" | find /i "node.exe" >nul
if errorlevel 1 (
    start "" wscript "start-background-server.vbs"
    timeout /t 1 /nobreak >nul
)
start http://localhost:3000
exit

