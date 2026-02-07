@echo off
setlocal

echo Creating/using venv...
if not exist .venv (
  python -m venv .venv
)

call .venv\Scripts\activate
if errorlevel 1 (
  echo Failed to activate venv.
  pause
  exit /b 1
)

echo Installing dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
echo.
echo Starting Druck Manager...
python druckmgr.py
pause
