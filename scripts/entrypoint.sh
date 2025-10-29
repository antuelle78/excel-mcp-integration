#!/bin/bash
# Activate virtual environment and run the application

cd /app
source venv/bin/activate
exec python src/main.py