@echo off
cd /d "%~dp0"
py -3 oss_photo_downloader.py
if errorlevel 1 pause
