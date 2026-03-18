@echo off
cd /d "%~dp0"

REM ── Check for SciLexer.dll ───────────────────────────────────────────────
if not exist "SciLexer.dll" (
    echo SciLexer.dll not found in this folder.
    echo.
    echo Trying to copy from Notepad++ installation...
    if exist "C:\Program Files\Notepad++\SciLexer.dll" (
        copy "C:\Program Files\Notepad++\SciLexer.dll" "SciLexer.dll"
        echo Copied from Notepad++.
    ) else (
        echo.
        echo SciLexer.dll not found in Notepad++ folder either.
        echo The app will fall back to the standard tk.Text editor.
        echo.
        echo To get Scintilla highlighting:
        echo   1. Download npp.7.9.5.portable.x64.zip from:
        echo      https://github.com/notepad-plus-plus/notepad-plus-plus/releases/tag/v7.9.5
        echo   2. Extract SciLexer.dll and place it in this folder.
        echo.
        pause
    )
)

REM ── Run the app ──────────────────────────────────────────────────────────
python wildcard_editor.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo Error running wildcard_editor.py
    echo Make sure Python is installed and pywin32 is available:
    echo   pip install pywin32
    pause
)
