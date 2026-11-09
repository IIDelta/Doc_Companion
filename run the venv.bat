@echo off
REM Navigate to the project directory
cd /d "%~dp0"

REM Activate the virtual environment
if exist .venv\Scripts\activate (
    call .venv\Scripts\activate
) else if exist venv\Scripts\activate (
    call venv\Scripts\activate
)

REM Run the application
python main.py
cmd /k
