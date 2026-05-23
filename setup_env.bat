@echo off
setlocal

echo [INFO] Starting setup for self-contained Python environment...

REM --- Configuration ---
set "PYTHON_VERSION=3.10.11"
set "PYTHON_DIR=python-embed"
set "PYTHON_ZIP=python-%PYTHON_VERSION%-embed-amd64.zip"
set "PYTHON_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/%PYTHON_ZIP%"
set "PTH_FILE=%PYTHON_DIR%\python310._pth"

REM --- 1. Check if Python folder already exists ---
if exist "%PYTHON_DIR%\python.exe" (
    echo [INFO] Python environment already exists. Verifying packages...
    goto :install_packages
)

REM --- 2. Download Embeddable Python ---
echo [INFO] Downloading Python %PYTHON_VERSION%...
powershell -Command "Invoke-WebRequest -Uri %PYTHON_URL% -OutFile %PYTHON_ZIP%"
if %errorlevel% neq 0 (
    echo [ERROR] Failed to download Python.
    goto :eof
)

echo [INFO] Extracting Python...
powershell -Command "Expand-Archive -Path '%PYTHON_ZIP%' -DestinationPath '%PYTHON_DIR%' -Force"
if %errorlevel% neq 0 (
    echo [ERROR] Failed to extract Python.
    goto :cleanup
)

REM --- 3. CRITICAL FIX: Enable site-packages for pip ---
echo [INFO] Enabling package installation...
powershell -Command "(Get-Content %PTH_FILE%) -replace '#import site', 'import site' | Set-Content %PTH_FILE%"
if %errorlevel% neq 0 (
    echo [ERROR] Failed to enable site-packages in Python environment.
    goto :cleanup
)

REM --- 4. Install Pip ---
echo [INFO] Downloading pip installer...
powershell -Command "Invoke-WebRequest -Uri https://bootstrap.pypa.io/get-pip.py -OutFile get-pip.py"
echo [INFO] Installing pip...
"%PYTHON_DIR%\python.exe" get-pip.py
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install pip.
    goto :cleanup
)

:install_packages
REM --- 5. Install Dependencies ---
echo [INFO] Installing required packages from requirements.txt...
"%PYTHON_DIR%\Scripts\pip.exe" install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install packages. Please check 'requirements.txt'.
    goto :cleanup
)

echo.
echo [SUCCESS] Setup is complete! The portable environment is ready.
echo You can now run the application using 'run_app.bat'.
echo.

:cleanup
REM --- 6. Clean up downloaded files ---
if exist "%PYTHON_ZIP%" del "%PYTHON_ZIP%"
if exist "get-pip.py" del "get-pip.py"

:eof
pause
