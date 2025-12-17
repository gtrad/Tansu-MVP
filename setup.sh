#!/bin/bash
# Tansu Setup Script - Run this once before first use

echo "Setting up Tansu..."

# Check for Python 3
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed."
    echo "Please install Python 3 from https://www.python.org/downloads/"
    exit 1
fi

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate and install dependencies
echo "Installing dependencies..."
source venv/bin/activate
pip install --upgrade pip
pip install customtkinter rumps

echo ""
echo "Setup complete! Run ./run.sh to start Tansu."
