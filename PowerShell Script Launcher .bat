@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul 2>&1

:: ============================================================================
:: PowerShell Script Launcher - Einfach aber robust
:: ============================================================================

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"
title PowerShell Script Launcher

:: PowerShell Version prüfen
pwsh --version >nul 2>&1
if %errorlevel% equ 0 (
    set "PS_EXE=pwsh"
    set "PS_VER=PowerShell 7+"
) else (
    set "PS_EXE=powershell"
    set "PS_VER=Windows PowerShell"
)

:menu
cls
echo.
echo ================================================================================
echo PowerShell Script Launcher - 2025 Jörn Walter - https://www.der-windows-papst.de
echo ================================================================================
echo.
echo Version: %PS_VER%
echo Verzeichnis: %SCRIPT_DIR%
echo.

:: PS1 Dateien auflisten und in Array speichern
set /a count=0
for %%f in ("*.ps1") do (
    if exist "%%f" (
        set /a count+=1
        set "script_!count!=%%f"
        echo [!count!] %%f
    )
)

if %count% equ 0 (
    echo Keine .ps1 Dateien im aktuellen Verzeichnis gefunden!
    echo.
    echo [R] Aktualisieren
    echo [X] Beenden
    echo.
    set /p choice="Wahl: "
    if /i "!choice!"=="R" goto menu
    if /i "!choice!"=="X" goto ende
    goto menu
)

echo.
echo [R] Aktualisieren   [X] Beenden
echo.

set /p auswahl="Script wählen (1-%count%, R, X): "

:: Sonderzeichen prüfen
if /i "%auswahl%"=="R" goto menu
if /i "%auswahl%"=="X" goto ende

:: Zahl validieren
if "%auswahl%" lss "1" goto ungueltig
if "%auswahl%" gtr "%count%" goto ungueltig

:: Script-Namen aus Array holen
set "script_name=!script_%auswahl%!"

echo.
echo ===============================================================================
echo Starte: !script_name!
echo ===============================================================================
echo.

:: Prüfen ob Datei existiert
if not exist "!script_name!" (
    echo FEHLER: Datei "!script_name!" nicht gefunden!
    pause
    goto menu
)

:: Script ausführen - ohne Profil um Konflikte zu vermeiden
echo Führe aus: %PS_EXE% -ExecutionPolicy Bypass -NoProfile -File "!script_name!"
echo.
%PS_EXE% -ExecutionPolicy Bypass -NoProfile -File "!script_name!"

set "exit_code=%errorlevel%"

echo.
echo ===============================================================================
if %exit_code% equ 0 (
    echo Script erfolgreich beendet
) else (
    echo Script beendet mit Exit-Code: %exit_code%
)
echo ===============================================================================
echo.

echo [1] Nochmal ausführen
echo [2] Script bearbeiten  
echo [3] Zurück zum Menü
echo [4] Beenden
echo.
set /p next="Wahl (1-4): "

if "%next%"=="1" (
    echo.
    echo Starte "!script_name!" erneut...
    echo.
    %PS_EXE% -ExecutionPolicy Bypass -NoProfile -File "!script_name!"
    echo.
    pause
    goto menu
)
if "%next%"=="2" (
    echo Öffne Editor...
    start "" notepad "!script_name!"
    goto menu
)
if "%next%"=="3" goto menu
if "%next%"=="4" goto ende

goto menu

:ungueltig
echo.
echo Ungültige Eingabe: %auswahl%
timeout 2 >nul
goto menu

:ende
echo.
echo Auf Wiedersehen!
pause
exit /b 0