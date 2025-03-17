@echo off
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -NoExit -File ins.ps1
