@echo off
:: Lancer le script PowerShell
powershell -ExecutionPolicy Bypass -File "%~dp0Harmoniser.ps1"
pause

:: Lancer GitHub Desktop
start "" "C:\Users\%USERNAME%\AppData\Local\GitHubDesktop\GitHubDesktop.exe"


