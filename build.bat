@echo off
echo Installing dependencies...
pip install -r requirements.txt

echo Building executable...
pyinstaller order_automation.spec

echo Build complete! Executable is in dist\OrderAutomation.exe
pause