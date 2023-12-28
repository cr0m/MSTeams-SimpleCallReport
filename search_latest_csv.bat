@echo off
setlocal enabledelayedexpansion

REM Find the latest .csv file
set "latest="
for /f "delims=" %%i in ('dir /b /od *.csv') do set latest=%%i

REM Check if a CSV file was found
if not defined latest (
    echo No CSV file found.
    exit /b
)

REM Run the Python script with the found CSV file
python search.py !latest! -inout -talktime
REM pause