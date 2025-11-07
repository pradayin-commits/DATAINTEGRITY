@echo off
setlocal ENABLEDELAYEDEXPANSION

cd /d "%~dp0"

echo.
echo [1/6] Checking Python...
where py >nul 2>nul && set PY=py || where python >nul 2>nul && set PY=python
if not defined PY (
  echo ERROR: Python not found.
  pause & exit /b 1
)

echo Using: %PY%
echo.

echo [2/6] Creating virtual environment...
if not exist ".venv" %PY% -m venv .venv
call .venv\Scripts\activate.bat || (echo Failed to activate venv & pause & exit /b 1)

echo.
echo [3/6] Upgrading pip...
python -m pip install --upgrade pip

echo.
echo [4/6] Installing requirements...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [5/6] CLEANING OLD BUILDS...
rmdir /s /q build dist 2>nul
:: DO NOT DELETE DIRT.spec

echo.
echo [6/6] BUILDING DIRT.EXE...
pyinstaller DIRT.spec --log-level=INFO
if %ERRORLEVEL% neq 0 (
  echo.
  echo BUILD FAILED
  pause & exit /b 1
)

echo.
echo SUCCESS! Opening folder...
explorer.exe "dist\DIRT"

echo.
echo Double-click DIRT.exe to run.
echo.
pause
endlocal