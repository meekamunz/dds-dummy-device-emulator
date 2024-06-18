@echo off

REM Clean build directories
rmdir /s /q __pycache__
rmdir /s /q build
rmdir /s /q dist
del *.spec

REM Compile using PyInstaller
pyinstaller --onefile --noconsole --icon=icon.ico DDS-config-builder.py --name "DDS Dummy Device Creator"

REM Pause to see the output
pause
