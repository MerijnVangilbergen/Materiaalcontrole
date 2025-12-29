@echo off
setlocal

REM --- Activate environment ---
echo Activating virtual environment...
call "venv\Scripts\activate.bat"
if %errorlevel% neq 0 (
    echo ERROR: Failed to activate the existing virtual environment.
    goto :end
)
echo Virtual environment activated.
echo.

REM --- Run the Python script ---
echo Starting the Python file Archive_Sanctions.py...
echo.
python Archive_Sanctions.py

REM Deactivate after the script finishes
call "venv\Scripts\deactivate.bat"

REM Check the exit code of the python script
echo.
if %errorlevel% neq 0 (
    echo The Python file closed unexpectedly with an error.
) else (
    echo The Python file finished successfully.
    )

echo Press any key to exit.
pause >nul