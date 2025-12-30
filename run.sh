#!/bin/bash
# Run script for Research Data Collection System

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "Virtual environment not found!"
    echo "Please run: ./setup.sh"
    exit 1
fi

# Activate virtual environment
source venv/bin/activate

# Run the application
echo "Starting Research Data Collection System..."
python run.py

