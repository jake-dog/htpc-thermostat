@echo off

rem Bypass "powershell security", and download hidapi/openhardwaremonitor
powershell.exe -ExecutionPolicy Bypass -File .\build.ps1
