@echo off
rem Excel tools: open the tool window (no PowerShell knowledge needed).
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0Start-ExcelTools.ps1" %*
