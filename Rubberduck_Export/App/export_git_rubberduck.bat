@echo off
chcp 65001 > nul
setlocal

:: === CONFIGURATION ===
set EXPORT_PATH=C:\VBA\GC_FISCALITÉ\PROJET_GIT\Rubberduck_Export
set SCRIPT_HARMONISE=%~dp0harmonise-casse-vba.ps1
set SCRIPT_UNICODE=%~dp0normaliser-accents-unicode.ps1

echo.
echo ➤ Vérification du dossier d’export...
if not exist "%EXPORT_PATH%" (
    echo ❌ Le dossier n’existe pas : %EXPORT_PATH%
    pause
    exit /b
)

echo.
echo ➤ Harmonisation de la casse des propriétés...
powershell.exe -ExecutionPolicy Bypass -File "%SCRIPT_HARMONISE%" -Root "%EXPORT_PATH%"

echo.
echo ➤ Normalisation des accents Unicode...
powershell.exe -ExecutionPolicy Bypass -File "%SCRIPT_UNICODE%" -Root "%EXPORT_PATH%"

echo.
echo ➤ Lancement de GitHub Desktop...
start "" "C:\Users\%USERNAME%\AppData\Local\GitHubDesktop\GitHubDesktop.exe"

echo.
echo ✅ Export, nettoyage et lancement terminés.
pause
