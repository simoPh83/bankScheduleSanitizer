#!/bin/bash

# Bank Schedule Sanitizer Launcher Script
# This script sets up the environment and launches the application

echo "🏦 Bank Schedule Sanitizer"
echo "=========================="

# Change to the script directory
cd "$(dirname "$0")"

# Check if virtual environment exists
if [ ! -d ".venv" ]; then
    echo "❌ Virtual environment not found. Setting up..."
    echo "Creating virtual environment..."
    python3 -m venv .venv
    
    echo "Activating virtual environment..."
    source .venv/bin/activate
    
    echo "Installing dependencies..."
    pip install -r requirements.txt
else
    echo "✅ Virtual environment found. Activating..."
    source .venv/bin/activate
fi

# Launch the application
echo "🚀 Launching Bank Schedule Sanitizer..."
python bank_schedule_sanitizer.py

echo "👋 Application closed."
