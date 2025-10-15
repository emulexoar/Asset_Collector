@echo off
REM Batch file to run the IT Asset Collection PowerShell script
REM Author: Marvin De Los Angeles, CIT Automation / AI / UX

powershell -ExecutionPolicy -File "%~dp0asset_collect.ps1"
IF %ERRORLEVEL% EQU 0 (
    DEL /F /Q "%~dp0asset_report.csv"
    DEL /F /Q "%~dp0asset_collect.ps1"
)