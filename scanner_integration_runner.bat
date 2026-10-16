@echo off
setlocal
cd /d "%~dp0"

rem The validated scanner core is a Python 3.8/S3PI workflow.
where py >nul 2>nul
if not errorlevel 1 (
    py -3.8 -c "import sys; print(sys.version_info[:2])" >nul 2>nul
    if not errorlevel 1 (
        py -3.8 -u "%~dp0scanner_integration_runner.py" %*
        exit /b %errorlevel%
    )
)

where pymanager >nul 2>nul
if not errorlevel 1 (
    pymanager exec -V:3.8 -c "import sys; print(sys.version_info[:2])" >nul 2>nul
    if not errorlevel 1 (
        pymanager exec -V:3.8 -u "%~dp0scanner_integration_runner.py" %*
        exit /b %errorlevel%
    )
)

where python3.8 >nul 2>nul
if not errorlevel 1 (
    python3.8 -u "%~dp0scanner_integration_runner.py" %*
    exit /b %errorlevel%
)

echo ERROR: Python 3.8 was not found for the scanner core.
echo Install/restore Python 3.8 or make it available through py -3.8, pymanager, or python3.8.
exit /b 38
