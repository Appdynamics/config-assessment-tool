@echo off
SETLOCAL

REM 1) Create venv if it doesn't exist
IF NOT EXIST .venv (
    echo Creating virtual environment...
    py -3 -m venv .venv
)

REM 2) Activate venv
CALL .venv\Scripts\activate.bat

REM 3) Install deps
echo Installing requirements...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

REM 4) Run the app
echo Starting Config Assessment Tool on http://127.0.0.1:5000 ...
python app.py

ENDLOCAL
