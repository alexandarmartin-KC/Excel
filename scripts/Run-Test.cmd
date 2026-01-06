@echo off
REM ============================================================================
REM Run-Test.cmd
REM ============================================================================
REM Beskrivelse:
REM   Double-click runner til at starte kamera-test manuelt.
REM   Ingen automatisk kørsel - kræver brugerinteraktion.
REM
REM Brug:
REM   Dobbeltklik på denne fil for at teste alle kameraer.
REM ============================================================================

echo.
echo ========================================
echo   KAMERA DASHBOARD - TEST RUNNER
echo ========================================
echo.
echo Dette script vil teste alle kameraer i Excel-filen.
echo.
echo BEMÆRK:
echo - Sørg for at CameraDashboard.xlsx er oprettet
echo - Test kører IKKE automatisk - du har startet det manuelt
echo - Efter test skal du manuelt refreshe Excel
echo.
pause
echo.

REM Tjek om PowerShell er tilgængeligt
where powershell >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo FEJL: PowerShell er ikke fundet på systemet!
    echo.
    pause
    exit /b 1
)

REM Kør PowerShell scriptet
echo Starter PowerShell script...
echo.
powershell.exe -ExecutionPolicy Bypass -File "%~dp0Check-Cameras.ps1"

REM Tjek exit code
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo   TEST GENNEMFØRT
    echo ========================================
) else (
    echo.
    echo ========================================
    echo   TEST FEJLEDE
    echo ========================================
    echo Tjek fejlmeddelelserne ovenfor.
)

echo.
pause
