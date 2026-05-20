@echo off
echo Starting ETF Updater...
docker compose up -d --build
echo Waiting for app to start...
timeout /t 6 /nobreak > nul
start http://localhost:8501
echo Done! The app should open in your browser.
echo To stop the app later, double-click stop.bat
