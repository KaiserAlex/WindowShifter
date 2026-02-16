@echo off
:: WindowShifter Launcher - runs the PowerShell script hidden
:: Prefers PowerShell 7 (pwsh) over Windows PowerShell 5.1 (powershell)
where pwsh >nul 2>&1
if %errorlevel% equ 0 (
    start "" /B pwsh.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0WindowShifter.ps1"
) else (
    start "" /B powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0WindowShifter.ps1"
)
