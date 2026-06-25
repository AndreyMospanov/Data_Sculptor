@echo off
cd /d "%~dp0"
python -m core.gui
if errorlevel 1 (
  echo.
  echo Data Sculptor failed to start.
  echo Try installing requirements:
  echo python -m pip install -r requirements.txt
  echo.
  pause
)
