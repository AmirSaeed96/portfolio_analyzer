#!/bin/bash
cd "$(dirname "$0")"
echo "Starting ETF Updater..."
docker compose up -d --build
echo "Waiting for app to start..."
sleep 6
open http://localhost:8501
echo "Done! The app should open in your browser."
echo "To stop the app later, run stop.sh"
