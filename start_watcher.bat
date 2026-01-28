@echo off
REM Word-to-LaTeX Thesis Converter Launcher
REM Starts the file watcher for automatic Word-to-LaTeX conversion

title Word-to-LaTeX Watcher

cd /d "C:\Users\dibis\OneDrive\Desktop\Thesis\LaTex\wi-thesis-template"

echo ============================================================
echo Word-to-LaTeX Thesis Converter
echo ============================================================
echo.
echo Watching: C:\Users\dibis\OneDrive\Desktop\Thesis\Manuscripts\Thesis_draft_chapter1_3.docx
echo Output:   C:\Users\dibis\OneDrive\Desktop\Thesis\LaTex\wi-thesis-template\chapters\
echo.
echo Press Ctrl+C to stop watching.
echo ============================================================
echo.

".venv\Scripts\python.exe" "scripts\word_to_latex.py"

pause
