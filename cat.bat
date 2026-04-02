@echo off
setlocal

REM ==========================================
REM Configuration
REM ==========================================
set REPO=ghcr.io/appdynamics/config-assessment-tool-windows

if exist VERSION (
    set /p VERSION=<VERSION
) else (
    echo Error: VERSION file not found.
    exit /b 1
)

set IMAGE=%REPO%:%VERSION%
set PORT=8501
set LOG_DIR=logs
set LOG_FILE=%LOG_DIR%\config-assessment-tool.log
set CONTAINER_NAME=cat-tool-container
set FILE_HANDLER_HOST=host.docker.internal

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

REM Define Mounts for Windows (Using %cd% for current directory)
set MOUNTS=-v "%cd%\input\jobs:C:\app\input\jobs" -v "%cd%\input\thresholds:C:\app\input\thresholds" -v "%cd%\output\archive:C:\app\output\archive" -v "%cd%\logs:C:\app\logs"

REM ==========================================
REM Argument Parsing
REM ==========================================
if "%1"=="--start" goto start
if "%1"=="--plugin" goto plugin
if "%1"=="shutdown" goto shutdown
goto usage

:start
shift
if "%1"=="docker" goto start_docker
goto start_local

REM ==========================================
REM Docker Startup Mode
REM ==========================================
:start_docker
REM 1. Start FileHandler on Host (Using python directly, NO pipenv)
if not exist "frontend\FileHandler.py" (
    echo Error: frontend\FileHandler.py not found.
    exit /b 1
)

echo Starting FileHandler service on host...
REM Kill any previous instance (rudimentary check by window title or just blindly)
taskkill /F /FI "WINDOWTITLE eq CAT-FileHandler" >nul 2>&1

REM Start in background with a specific title so we can kill it later
start "CAT-FileHandler" /min cmd /c "python frontend\FileHandler.py >> %LOG_FILE% 2>&1"
echo FileHandler started.
timeout /t 2 /nobreak >nul

REM 2. Reset Container
docker stop %CONTAINER_NAME% >nul 2>&1
docker rm %CONTAINER_NAME% >nul 2>&1

shift
REM Check if there are additional arguments (Backend vs UI)
if "%1"=="" (
    echo Starting container in UI mode...
    REM Note: We do NOT pass "streamlit run..." args here.
    REM The Dockerfile ENTRYPOINT defaults to frontend if no "backend" arg is present.
    REM We pass --server.headless=true to ensure non-interactive start if needed,
    REM though the Entrypoint handles 'python -m streamlit run ...' automatically.

    docker run -d --name %CONTAINER_NAME% -e FILE_HANDLER_HOST=%FILE_HANDLER_HOST% -p %PORT%:%PORT% %MOUNTS% %IMAGE% --server.headless=true

    if %ERRORLEVEL% EQU 0 (
        echo Container started successfully.
        echo UI available at http://localhost:%PORT%
        docker logs -f %CONTAINER_NAME%
    ) else (
        echo Failed to start container.
        exit /b 1
    )
) else (
    echo Starting container in backend mode with args: %1 %2 %3 %4 %5 %6 %7 %8 %9
    REM Pass "backend" + user arguments to trigger the backend branch in entrypoint.bat
    docker run --rm --name %CONTAINER_NAME% -e FILE_HANDLER_HOST=%FILE_HANDLER_HOST% -p %PORT%:%PORT% %MOUNTS% %IMAGE% backend %1 %2 %3 %4 %5 %6 %7 %8 %9

    if %ERRORLEVEL% EQU 0 (
        echo Container finished.
    ) else (
        echo Failed to start container.
        exit /b 1
    )
)
goto end


REM ==========================================
REM Local Source Startup Mode
REM ==========================================
:start_local
REM Setting PYTHONPATH for local execution
set PYTHONPATH=%cd%;%cd%\backend

REM Check for pipenv
set PIPENV_CMD=pipenv
pipenv --version >nul 2>&1
if %ERRORLEVEL% EQU 0 goto pipenv_found

python -m pipenv --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PIPENV_CMD=python -m pipenv
    goto pipenv_found
)

echo pipenv not found. Attempting to install via pip...
pip install pipenv
if %ERRORLEVEL% NEQ 0 (
    echo Error: Failed to install pipenv. Please ensure Python and pip are installed and in your PATH.
    exit /b 1
)

REM Check again after install
python -m pipenv --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PIPENV_CMD=python -m pipenv
    goto pipenv_found
)
set PIPENV_CMD=pipenv

:pipenv_found
REM Ensure dependencies are installed
echo Checking/Installing dependencies...
%PIPENV_CMD% install

if "%1"=="" (
    echo Running application in UI mode from source...
    echo UI available at http://localhost:%PORT%
    %PIPENV_CMD% run streamlit run frontend\frontend.py
) else (
    echo Running application in backend mode from source with args: %1 %2 %3 %4 %5 %6 %7 %8 %9
    %PIPENV_CMD% run python backend\backend.py %1 %2 %3 %4 %5 %6 %7 %8 %9
)
goto end


REM ==========================================
REM Plugin Management
REM ==========================================
:plugin
REM Pipenv detection for plugins
set PIPENV_CMD=pipenv
pipenv --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    python -m pipenv --version >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        set PIPENV_CMD=python -m pipenv
    ) else (
         echo pipenv not found. Please run --start to install dependencies first.
         exit /b 1
    )
)

shift
if "%1"=="list" goto plugin_list
if "%1"=="docs" goto plugin_docs
if "%1"=="start" goto plugin_start
echo Error: Unknown plugin command.
goto usage

:plugin_list
set PYTHONPATH=%cd%;%cd%\backend
%PIPENV_CMD% run python backend\plugin_manager.py list
goto end

:plugin_docs
shift
if "%1"=="" (
    echo Error: Plugin name required.
    exit /b 1
)
set PLUGIN_NAME=%1
set PYTHONPATH=%cd%;%cd%\backend
%PIPENV_CMD% run python backend\plugin_manager.py docs %PLUGIN_NAME%
goto end

:plugin_start
shift
if "%1"=="" (
    echo Error: Plugin name required.
    exit /b 1
)
set PLUGIN_NAME=%1
shift
REM Capture arguments manually for simplicity in batch
set ARGS=%1 %2 %3 %4 %5 %6 %7 %8 %9
set PYTHONPATH=%cd%;%cd%\backend
%PIPENV_CMD% run python backend\plugin_manager.py start %PLUGIN_NAME% %ARGS%
goto end


REM ==========================================
REM Shutdown Mode
REM ==========================================
:shutdown
echo Shutting down container: %CONTAINER_NAME%
docker stop %CONTAINER_NAME% >nul 2>&1
docker rm %CONTAINER_NAME% >nul 2>&1
echo Container stopped and removed.

echo Stopping FileHandler process...
taskkill /F /FI "WINDOWTITLE eq CAT-FileHandler" >nul 2>&1
echo FileHandler stopped.

echo Stopping generic python processes (backend/streamlit)...
REM Warning: This might be too aggressive on a developer machine,
REM but matches the aggressive pkill in config-assessment-tool.sh
taskkill /F /IM python.exe /T >nul 2>&1
taskkill /F /IM streamlit.exe /T >nul 2>&1
echo Processes stopped.
goto end


REM ==========================================
REM Usage / Help
REM ==========================================
:usage
echo Usage:
echo   cat.bat --start                # Starts CAT UI. Requires Python 3.12 and pipenv installed. UI accessible at http://localhost:8501
echo   cat.bat --start [args]         # Starts CAT headless mode from source with [args]. Requires Python 3.12 ^& pipenv installed.
echo   cat.bat --start docker         # Starts CAT UI using Docker. Requires Docker. UI accessible at http://localhost:8501
echo   cat.bat --start docker [args]  # Starts CAT headless mode using Docker with [args]. Requires Docker installed.
echo   cat.bat --plugin ^<list^|start^|docs^> [name]    # list plugins ^| start plugin ^| show docs for plugin
echo   cat.bat shutdown               # Stop container and processes
echo.
echo Arguments [args]:
echo   -j, --job-file ^<name^>             Job file name (default: DefaultJob)
echo   -t, --thresholds-file ^<name^>      Thresholds file name (default: DefaultThresholds)
echo   -d, --debug                       Enable debug logging
echo   -c, --concurrent-connections ^<n^>  Number of concurrent connections
goto end

:end
endlocal