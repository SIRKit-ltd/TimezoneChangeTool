@echo off
setlocal EnableExtensions DisableDelayedExpansion
title Timezone Change Tool
echo Preparing Timezone Change Tool...
echo.
if not exist "%~dp0Setup-TimezoneChangeTool.ps1" goto MissingFiles
if not exist "%~dp0Start-TimezoneChangeTool.ps1" goto MissingFiles
if not exist "%~dp0TimezoneEngine.psm1" goto MissingFiles
if not exist "%~dp0SeriesMigration.ps1" goto MissingFiles
if not exist "%~dp0GraphRequestContract.ps1" goto MissingFiles
if not exist "%~dp0config.json" goto MissingFiles
set "TZ_PWSH="
if defined ProgramW6432 if exist "%ProgramW6432%\PowerShell\7\pwsh.exe" set "TZ_PWSH=%ProgramW6432%\PowerShell\7\pwsh.exe"
if not defined TZ_PWSH if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" set "TZ_PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not defined TZ_PWSH for /f "delims=" %%P in ('where pwsh.exe 2^>nul') do if not defined TZ_PWSH set "TZ_PWSH=%%P"
if not defined TZ_PWSH goto MissingPowerShell
pushd "%~dp0"
if errorlevel 1 goto FolderError
"%TZ_PWSH%" -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0Setup-TimezoneChangeTool.ps1" %*
set "TZ_EXIT=%ERRORLEVEL%"
popd
if "%TZ_EXIT%"=="0" exit /b 0
echo.
echo The tool could not start or stopped with an error. Exit code: %TZ_EXIT%
echo Copy the error above before closing this window.
echo If Windows says the scripts are blocked, ask IT to review and unblock
echo the downloaded files. Organization security policies still apply.
pause
exit /b %TZ_EXIT%

:MissingFiles
echo Required files are missing beside this launcher.
echo Extract the entire ZIP first, then run the launcher from that folder.
pause
exit /b 2

:MissingPowerShell
echo PowerShell 7.2 or later is required and could not be found.
echo Ask IT to install PowerShell 7, then reopen this launcher.
pause
exit /b 3

:FolderError
echo The extracted tool folder could not be opened.
echo Extract the ZIP to a local folder and try again.
pause
exit /b 4
