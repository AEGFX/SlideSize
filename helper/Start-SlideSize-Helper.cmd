@echo off
rem SlideSize PowerPoint to PDF helper. Double-click to start. Nothing is installed.
title SlideSize PowerPoint to PDF helper
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0slidesize-helper.ps1"
if errorlevel 1 pause
