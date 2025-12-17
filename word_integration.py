"""
Word integration module - OS-detecting wrapper.
Imports the appropriate platform-specific module based on the current OS.
"""

import platform

# Import the appropriate module based on OS
if platform.system() == "Windows":
    from word_windows import WordIntegration, DocumentInfo, HAS_WIN32 as HAS_WORD
elif platform.system() == "Darwin":
    from word_mac import WordIntegration, DocumentInfo, HAS_APPLESCRIPT as HAS_WORD
else:
    WordIntegration = None
    DocumentInfo = None
    HAS_WORD = False

__all__ = ['WordIntegration', 'DocumentInfo', 'HAS_WORD']
