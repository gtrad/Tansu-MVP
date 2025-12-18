# -*- mode: python ; coding: utf-8 -*-
# Tansu - Cross-platform PyInstaller spec file
# Builds for Mac (.app) or Windows (.exe) depending on current platform

from PyInstaller.utils.hooks import collect_all
import sys
import os

# Detect platform
is_mac = sys.platform == 'darwin'
is_windows = sys.platform == 'win32'

# Base data files (all platforms)
datas = [
    ('database.py', '.'),
    ('word_integration.py', '.'),
    ('excel_reader.py', '.'),
    ('docx_updater.py', '.'),
    ('api_server.py', '.'),
    ('version.py', '.'),
    ('settings.py', '.'),
    ('update_checker.py', '.'),
    ('icon.png', '.'),  # App icon for window
]

binaries = []

# Base hidden imports
hiddenimports = [
    'customtkinter',
    'darkdetect',
    'openpyxl',
    'docx',
    'sqlite3',
]

# Collect CustomTkinter data (required for themes)
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]
binaries += tmp_ret[1]
hiddenimports += tmp_ret[2]

# Global hotkey support (pynput) - both platforms
hiddenimports += ['pynput', 'pynput.keyboard', 'pynput.keyboard._darwin', 'pynput.keyboard._win32']
try:
    tmp_ret = collect_all('pynput')
    datas += tmp_ret[0]
    binaries += tmp_ret[1]
    hiddenimports += tmp_ret[2]
except:
    pass

# Platform-specific configuration
if is_mac:
    # Mac-specific files
    datas += [
        ('word_mac.py', '.'),
        ('app.py', '.'),
    ]
    hiddenimports += ['objc', 'AppKit', 'Foundation', 'ApplicationServices', 'Quartz']

    entry_script = 'app.py'

elif is_windows:
    # Windows-specific files
    datas += [
        ('word_windows.py', '.'),
        ('app.py', '.'),
        ('word_addin', 'word_addin'),  # Include VBA files
    ]
    hiddenimports += ['win32com', 'win32com.client', 'PIL', 'PIL.Image']

    # Collect Pillow (needed for window icon)
    try:
        tmp_ret = collect_all('PIL')
        datas += tmp_ret[0]
        binaries += tmp_ret[1]
        hiddenimports += tmp_ret[2]
    except:
        pass

    entry_script = 'app.py'

else:
    entry_script = 'app.py'

# Analysis
a = Analysis(
    [entry_script],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'scipy', 'pandas', 'pytest'],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Tansu',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file='entitlements.plist' if is_mac else None,
    icon='icon.icns' if is_mac and os.path.exists('icon.icns') else ('icon.ico' if is_windows and os.path.exists('icon.ico') else None),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Tansu',
)

# Mac-specific: Create .app bundle
if is_mac:
    app = BUNDLE(
        coll,
        name='Tansu.app',
        icon='icon.icns' if os.path.exists('icon.icns') else None,
        bundle_identifier='com.tansu.variabletracker',
        info_plist={
            'NSPrincipalClass': 'NSApplication',
            'NSHighResolutionCapable': 'True',
            'LSMinimumSystemVersion': '10.13.0',
            'NSAppleEventsUsageDescription': 'Tansu needs to control Microsoft Word to insert and update variables.',
            'CFBundleShortVersionString': '0.9.0',
        },
    )
