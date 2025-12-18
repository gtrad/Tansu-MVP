#!/bin/bash
# Create a DMG with Applications folder shortcut for drag-and-drop install

set -e

APP_NAME="Tansu"
DMG_NAME="Tansu-0.9.0"
VOL_NAME="Tansu Installer"
APP_PATH="dist/Tansu.app"
DMG_FINAL="${DMG_NAME}.dmg"

# Check if app exists
if [ ! -d "$APP_PATH" ]; then
    echo "Error: $APP_PATH not found. Run ./build_mac.sh first."
    exit 1
fi

echo "Creating DMG installer..."

# Clean up any existing DMG
rm -f "$DMG_FINAL"

# Create a temporary directory for DMG contents
DMG_DIR="dist/dmg_contents"
rm -rf "$DMG_DIR"
mkdir -p "$DMG_DIR"

# Copy the app
cp -R "$APP_PATH" "$DMG_DIR/"

# Create Applications symlink
ln -s /Applications "$DMG_DIR/Applications"

# Create the DMG directly
hdiutil create -volname "$VOL_NAME" -srcfolder "$DMG_DIR" -ov -format UDZO "$DMG_FINAL"

# Clean up
rm -rf "$DMG_DIR"

echo ""
echo "==================================="
echo "DMG created: $DMG_FINAL"
echo "==================================="
echo ""
echo "Users can drag Tansu.app to Applications folder."
