#!/bin/bash
# Build Tansu.app for macOS

echo "==================================="
echo "Building Tansu for macOS"
echo "==================================="

# Check if we're on macOS
if [[ "$OSTYPE" != "darwin"* ]]; then
    echo "Error: This script must be run on macOS"
    exit 1
fi

# Get script directory
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Check for Python
if ! command -v python3 &> /dev/null; then
    echo "Error: python3 not found"
    exit 1
fi

# Create/activate virtual environment (optional but recommended)
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

echo "Activating virtual environment..."
source venv/bin/activate

# Install dependencies
echo "Installing dependencies..."
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller rumps

# Clean previous build
echo "Cleaning previous build..."
rm -rf build dist

# Build the app
echo "Building application..."
pyinstaller --noconfirm Tansu.spec

# Check if build succeeded
if [ -d "dist/Tansu.app" ]; then
    echo ""
    echo "==================================="
    echo "Build successful!"
    echo "==================================="
    echo "App location: dist/Tansu.app"
    echo ""
    echo "To run: open dist/Tansu.app"
    echo ""
    echo "To create a DMG for distribution:"
    echo "  hdiutil create -volname Tansu -srcfolder dist/Tansu.app -ov -format UDZO Tansu.dmg"
else
    echo ""
    echo "==================================="
    echo "Build failed!"
    echo "==================================="
    echo "Check the output above for errors."
    exit 1
fi
