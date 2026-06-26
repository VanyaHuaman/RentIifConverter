@echo off
setlocal

where pwsh.exe >nul 2>nul
if errorlevel 1 (
    echo Rent IIF Converter could not find PowerShell 7 ^(pwsh.exe^) on PATH.
    echo.
    echo Install PowerShell 7, then try again:
    echo   winget install --id Microsoft.PowerShell --source winget
    echo.
    pause
    exit /b 1
)

set "POWERSHELL_UPDATECHECK=Off"
pwsh.exe -STA -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0RentIifConverter.ps1"
exit /b %ERRORLEVEL%
