@echo off
setlocal

cd /d "%~dp0"

set "PYTHON=%~dp0.thinkcell\Scripts\python.exe"
set "APP=%~dp0.thinkcell\app.py"

if not exist "%PYTHON%" (
    echo.
    echo ERROR: Python was not found here:
    echo %PYTHON%
    echo.
    echo Make sure the .thinkcell environment is present in this project.
    pause
    exit /b 1
)

if not exist "%APP%" (
    echo.
    echo ERROR: Slide Maker app was not found here:
    echo %APP%
    echo.
    pause
    exit /b 1
)

echo Starting ThinkCell Slide Maker...
echo.

"%PYTHON%" -m streamlit run "%APP%"

set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo Slide Maker exited with error code %EXIT_CODE%.
    pause
)

exit /b %EXIT_CODE%
