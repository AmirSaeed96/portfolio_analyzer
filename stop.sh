#!/bin/bash
cd "$(dirname "$0")"
echo "Stopping ETF Updater..."
docker compose down
echo "Stopped."
