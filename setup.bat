@echo off
echo =============================================
echo   Account Reconciliation Tool - First Time Setup
echo =============================================
echo.

:: Backend setup
echo [1/3] Creating Python virtual environment...
cd flask-backend
python -m venv venv
echo Done.
echo.

echo [2/3] Installing Python dependencies...
call venv\Scripts\activate
pip install -r requirements.txt
echo Done.
echo.

:: Create .env if it doesn't exist
if not exist .env (
    copy .env.example .env
    echo Created flask-backend\.env from .env.example
    echo.
    echo  IMPORTANT: Open flask-backend\.env and fill in your API key before running the app.
    echo  - OPEN_ROUTER_API_KEY  (required for AI matching)
    echo.
    echo  Get a free key at: https://openrouter.ai/keys
    echo.
) else (
    echo flask-backend\.env already exists, skipping.
    echo.
)

cd ..

echo [3/3] Setup complete!
echo.
echo =============================================
echo   Next steps:
echo   1. Edit flask-backend\.env with your OpenRouter API key
echo   2. Double-click start.bat to launch the app
echo =============================================
pause
