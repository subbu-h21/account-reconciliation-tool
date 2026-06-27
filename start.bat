@echo off
echo Starting Account Reconciliation Tool...

:: Start Flask backend in a new window
start "Account Reconciliation" cmd /k "cd flask-backend && venv\Scripts\activate && python app.py"

:: Wait for server to be ready then open browser
timeout /t 3 /nobreak >nul
start "" "http://localhost:5000"
