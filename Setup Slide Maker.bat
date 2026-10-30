@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "ROOT=%~dp0"
set "ENV_DIR=%ROOT%.thinkcell"
set "APP=%ENV_DIR%\app.py"
set "REQ=%ROOT%requirements.txt"
set "ENV_PYTHON=%ENV_DIR%\Scripts\python.exe"

echo.
echo ============================================================
echo   ThinkCell Slide Maker - First-Time Setup
echo ============================================================
echo.

if not exist "%APP%" (
    echo ERROR: Could not find:
    echo   %APP%
    echo.
    echo Make sure this file is in the repository root,
    echo next to the .thinkcell folder.
    echo.
    pause
    exit /b 1
)

if not exist "%REQ%" (
    echo ERROR: Could not find:
    echo   %REQ%
    echo.
    pause
    exit /b 1
)

if exist "%ENV_PYTHON%" (
    echo Existing local Python environment found.
    echo It will be reused and dependencies will be refreshed.
    goto INSTALL
)

echo Creating the local Python environment inside .thinkcell...
echo Your source files in .thinkcell\App and .thinkcell\UI will be preserved.
echo.

where py >nul 2>&1
if not errorlevel 1 (
    py -3 -m venv "%ENV_DIR%"
    if errorlevel 1 goto VENV_FAILED
    goto VENV_READY
)

where python >nul 2>&1
if not errorlevel 1 (
    python -m venv "%ENV_DIR%"
    if errorlevel 1 goto VENV_FAILED
    goto VENV_READY
)

echo ERROR: Python was not found on this computer.
echo.
echo Install Python 3 for Windows, then run this setup again.
echo During installation, enable the option to add Python to PATH
echo or install the Python Launcher.
echo.
pause
exit /b 1

:VENV_FAILED
echo.
echo ERROR: The Python virtual environment could not be created.
echo Check that Python is installed correctly and try again.
echo.
pause
exit /b 1

:VENV_READY
rem Python 3.13+ may create a venv-specific .gitignore containing "*".
rem Since .thinkcell also contains this project's source code, remove only
rem the automatically generated venv ignore file if it is present.
if exist "%ENV_DIR%\.gitignore" (
    findstr /C:"Created by venv" "%ENV_DIR%\.gitignore" >nul 2>&1
    if not errorlevel 1 del /q "%ENV_DIR%\.gitignore" >nul 2>&1
)

:INSTALL
echo.
echo Upgrading pip...
"%ENV_PYTHON%" -m pip install --upgrade pip
if errorlevel 1 goto INSTALL_FAILED

echo.
echo Installing ThinkCell Slide Maker dependencies...
"%ENV_PYTHON%" -m pip install -r "%REQ%"
if errorlevel 1 goto INSTALL_FAILED

echo.
echo Verifying required Python packages...
"%ENV_PYTHON%" -c "import streamlit, openpyxl, win32com.client, pptx; print('Python package check passed.')"
if errorlevel 1 goto INSTALL_FAILED

echo.
echo ============================================================
echo   Setup complete.
echo ============================================================
echo.
echo You can now close this window and double-click:
echo   Run Slide Maker.bat
echo.
pause
exit /b 0

:INSTALL_FAILED
echo.
echo ERROR: Setup did not complete successfully.
echo Review the error messages above.
echo.
pause
exit /b 1
