#!/bin/bash
# Tansu Run Script

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Check if venv exists
if [ ! -d "venv" ]; then
    echo "Please run ./setup.sh first"
    exit 1
fi

# Activate virtual environment and run
source venv/bin/activate
python app.py
