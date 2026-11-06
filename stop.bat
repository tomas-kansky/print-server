@echo off
taskkill /F /IM pythonw.exe 2>nul
taskkill /F /IM python.exe 2>nul
echo Print Server stopped.
pause
