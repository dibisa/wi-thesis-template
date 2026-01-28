@echo off
REM Word-to-LaTeX Single Conversion
REM Converts Word document to LaTeX once (no watching)

cd /d "C:\Users\dibis\OneDrive\Desktop\Thesis\LaTex\wi-thesis-template"

echo Converting Word to LaTeX...
".venv\Scripts\python.exe" "scripts\word_to_latex.py" --once

echo.
echo Done!
pause
