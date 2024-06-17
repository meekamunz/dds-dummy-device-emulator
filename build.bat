:START
@echo off

:CLEANBUILD
rmdir /s /q __pycache__
rmdir /s /q build
rmdir /s /q dist
del *.spec

:PYTHONCOMPILE
pyinstaller --onefile --icon=icon.ico DDS-config-builder.py -n "DDS Dummy Device Creator"

REM copy .\dist\*.exe "C:\Users\MichaelMunns\OneDrive - GrassValley\Product Downloads\Toolkit\IPRA"