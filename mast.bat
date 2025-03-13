@echo off
"C:\Users\josej\OneDrive\Documents\Python Scripts\python.exe" mast.py
pause
exit
if %errorlevel% neq 0 (
    echo Error: Python script did not execute successfully.
)