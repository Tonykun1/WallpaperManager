@echo off
REM Start Wallpaper Engine Manager in background

cd /d "%~dp0"
start /B pythonw.exe wallpaper_engine_manager.py

exit
```