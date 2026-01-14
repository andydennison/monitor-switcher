@echo off
REM Start Monitor Switcher in background
REM Use this batch file to start the application easily

cd /d "%~dp0"
start /min pythonw monitor_switcher.py
