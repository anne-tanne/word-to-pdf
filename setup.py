"""
Build with:
    source venv/bin/activate
    python3 setup.py py2app
"""
from setuptools import setup

APP = ["main.py"]
DATA_FILES = ["locales"]
OPTIONS = {
    "argv_emulation": False,
    "iconfile": "AppIcon.icns",
    # Include tkinterdnd2 as real files so its native tkdnd library is bundled
    # and drag & drop works in the built app.
    "packages": ["tkinterdnd2"],
    "plist": {
        "CFBundleName": "Word to PDF",
        "CFBundleDisplayName": "Word to PDF",
        "CFBundleIdentifier": "local.wordtopdf.converter",
        "CFBundleShortVersionString": "1.0.0",
        "NSHighResolutionCapable": True,
        "LSMinimumSystemVersion": "11.0",
        "LSUIElement": False,
        "NSAppleEventsUsageDescription": "This app needs to control Microsoft Word to convert your documents to PDF.",
    },
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={"py2app": OPTIONS},
    setup_requires=["py2app"],
)
