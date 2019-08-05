@REM Launch Powershell script with same name as the batch file
@echo off
"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy Bypass -File %~dpn0.ps1
@REM pause