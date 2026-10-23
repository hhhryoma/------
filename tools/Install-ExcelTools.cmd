@echo off
rem Excel tools: install to this PC (or reinstall) from this folder and register the SendTo menu.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-ExcelTools.ps1" %*
pause
