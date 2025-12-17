"""
Setup script for building Tansu as a macOS application.
Run with: python setup.py py2app
"""

from setuptools import setup

APP = ['app.py']
DATA_FILES = [
    'menubar_app.py',
    'database.py',
    'word_mac.py',
    'word_integration.py',
]

OPTIONS = {
    'argv_emulation': False,
    'iconfile': None,  # Add icon later if desired
    'plist': {
        'CFBundleName': 'Tansu',
        'CFBundleDisplayName': 'Tansu',
        'CFBundleIdentifier': 'com.tansu.variabletracker',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '10.13.0',
        'NSAppleEventsUsageDescription': 'Tansu needs to control Microsoft Word to insert and update variables.',
    },
    'includes': [
        'customtkinter',
        'rumps',
        'tkinter',
        'sqlite3',
    ],
    'packages': [
        'customtkinter',
    ],
    'frameworks': [],
}

setup(
    name='Tansu',
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
