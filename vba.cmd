@echo off
setlocal
rem Prefer Windows PowerShell modules when launched from a PowerShell 7 terminal.
set "PSModulePath=%SystemRoot%\System32\WindowsPowerShell\v1.0\Modules;%PSModulePath%"
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0vba.ps1" %*
exit /b %ERRORLEVEL%
