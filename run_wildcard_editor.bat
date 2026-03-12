@echo off
:: Wildcard Editor Launcher
:: Ensures pyspellchecker is installed using the SAME Python that runs the app.

:: ── Find Python ──────────────────────────────────────────────────────────────
set PYTHON_EXE=python
%PYTHON_EXE% --version >nul 2>&1
if errorlevel 1 (
    set PYTHON_EXE=python3
    %PYTHON_EXE% --version >nul 2>&1
    if errorlevel 1 (
        set PYTHON_EXE=py
        %PYTHON_EXE% --version >nul 2>&1
        if errorlevel 1 (
            echo ERROR: Could not find Python. Install Python 3.8+ and add to PATH.
            pause
            exit /b 1
        )
    )
)

echo Using: %PYTHON_EXE%

:: ── Install pyspellchecker using the EXACT same Python we will run ───────────
%PYTHON_EXE% -c "import spellchecker" >nul 2>&1
if errorlevel 1 (
    echo Installing pyspellchecker...
    %PYTHON_EXE% -m pip install pyspellchecker --quiet
    if errorlevel 1 (
        echo WARNING: Could not install pyspellchecker. Spell check unavailable.
    ) else (
        echo pyspellchecker installed.
    )
)

:: ── Launch ───────────────────────────────────────────────────────────────────
set SCRIPT_DIR=%~dp0
%PYTHON_EXE% "%SCRIPT_DIR%wildcard_editor.py"
