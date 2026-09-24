@echo off
setlocal enabledelayedexpansion
set "LATEST_FILE="
for /f "delims=" %%F in ('dir /b /o:-d "%~dp0Masterlist 2027-2028_KQ_*.xlsx" 2^>nul') do (
    if not defined LATEST_FILE set "LATEST_FILE=%~dp0%%F"
)
if defined LATEST_FILE (
    start "" "!LATEST_FILE!"
) else (
    start "" "%~dp0Masterlist 2027-2028_Mau.xlsx"
)
