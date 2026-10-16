@echo off
setlocal
cd /d "%~dp0"
set "CAS_CURATOR_EDITION=PUBLIC"

echo CAS Curator v1.0.1
echo.

set "PYEXE="
set "VENV_PY=%CD%\.venv\Scripts\python.exe"

rem Prefer the new Python Install Manager. Keep the existing Sims scanner
rem Python 3.8 untouched; this GUI requires Python 3.10-3.14 64-bit.
where pymanager >nul 2>nul
if not errorlevel 1 (
  echo Checking Python 3.14 through Python Install Manager...
  pymanager exec -V:3.14 -c "import struct,sys; raise SystemExit(0 if sys.version_info[:2]==(3,14) and struct.calcsize('P')*8==64 else 1)" >nul 2>nul && set "PYEXE=pymanager exec -V:3.14"
  if not defined PYEXE pymanager exec -V:3.13 -c "import struct,sys; raise SystemExit(0 if sys.version_info[:2]==(3,13) and struct.calcsize('P')*8==64 else 1)" >nul 2>nul && set "PYEXE=pymanager exec -V:3.13"
  if not defined PYEXE pymanager exec -V:3.12 -c "import struct,sys; raise SystemExit(0 if sys.version_info[:2]==(3,12) and struct.calcsize('P')*8==64 else 1)" >nul 2>nul && set "PYEXE=pymanager exec -V:3.12"
  if not defined PYEXE pymanager exec -V:3.11 -c "import struct,sys; raise SystemExit(0 if sys.version_info[:2]==(3,11) and struct.calcsize('P')*8==64 else 1)" >nul 2>nul && set "PYEXE=pymanager exec -V:3.11"
  if not defined PYEXE pymanager exec -V:3.10 -c "import struct,sys; raise SystemExit(0 if sys.version_info[:2]==(3,10) and struct.calcsize('P')*8==64 else 1)" >nul 2>nul && set "PYEXE=pymanager exec -V:3.10"
)

rem Fallback for traditional Python installations / old launcher.
if not defined PYEXE py -3.14 -c "import struct; raise SystemExit(0 if struct.calcsize('P')*8 == 64 else 1)" >nul 2>nul && set "PYEXE=py -3.14"
if not defined PYEXE py -3.13 -c "import struct; raise SystemExit(0 if struct.calcsize('P')*8 == 64 else 1)" >nul 2>nul && set "PYEXE=py -3.13"
if not defined PYEXE py -3.12 -c "import struct; raise SystemExit(0 if struct.calcsize('P')*8 == 64 else 1)" >nul 2>nul && set "PYEXE=py -3.12"
if not defined PYEXE py -3.11 -c "import struct; raise SystemExit(0 if struct.calcsize('P')*8 == 64 else 1)" >nul 2>nul && set "PYEXE=py -3.11"
if not defined PYEXE py -3.10 -c "import struct; raise SystemExit(0 if struct.calcsize('P')*8 == 64 else 1)" >nul 2>nul && set "PYEXE=py -3.10"

rem Last fallback: global python, but only if it is a supported 64-bit version.
if not defined PYEXE python -c "import struct,sys; raise SystemExit(0 if (3,10) <= sys.version_info[:2] <= (3,14) and struct.calcsize('P')*8 == 64 else 1)" >nul 2>nul && set "PYEXE=python"

if not defined PYEXE (
  echo.
  echo Python 3.10-3.14 64-bit could not be launched.
  echo.
  echo If you just installed the new Python Install Manager, open Command Prompt and run:
  echo     pymanager exec -V:3.14 --version
  echo.
  echo Your existing Sims scanner Python 3.8 can remain installed unchanged.
  pause
  exit /b 1
)

echo Using: %PYEXE%

rem IMPORTANT: Never activate a copied venv and then call bare "python".
rem activate.bat contains an absolute path to the folder where the venv was
rem originally created. That can silently fall back to the user's Python 3.8.
rem Instead, always execute this folder's venv python by its full path.

if exist "%VENV_PY%" (
  "%VENV_PY%" -c "import struct,sys; raise SystemExit(0 if (3,10) <= sys.version_info[:2] <= (3,14) and struct.calcsize('P')*8 == 64 else 1)" >nul 2>nul
  if errorlevel 1 (
    echo Existing private .venv is stale or incompatible. Rebuilding it...
    rmdir /s /q ".venv" || goto :fail
  )
)

if not exist "%VENV_PY%" (
  echo Creating private CAS Curator environment...
  %PYEXE% -m venv ".venv" || goto :fail
)

"%VENV_PY%" -c "import PySide6; import sys; raise SystemExit(0)" >nul 2>nul
if errorlevel 1 (
  echo Installing PySide6 into this folder's private .venv...
  "%VENV_PY%" -m pip install --upgrade pip || goto :fail
  "%VENV_PY%" -m pip install -r requirements.txt || goto :fail
)

echo Starting CAS Curator...
"%VENV_PY%" app.py
if errorlevel 1 goto :appcrash
exit /b 0

:appcrash
echo.
echo CAS Curator stopped with an error.
echo Please copy or screenshot the error shown above and report it with the error details.
pause
exit /b 1

:fail
echo.
echo Could not start CAS Curator.
echo Copy everything shown above and report it with the error details so the launcher can be fixed.
pause
exit /b 1
