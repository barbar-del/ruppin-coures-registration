@echo off
setlocal enabledelayedexpansion

REM Activate the virtual environment
call .venv\Scripts\activate
pip install -r requirements.txt

REM Run your Python application
python main.py

REM Deactivate the virtual environment
call .venv\Scripts\deactivate
pause