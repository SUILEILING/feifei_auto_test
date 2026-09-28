@echo off
REM ============================================================
REM  Build SourenTestTool
REM  IMPORTANT: Always build with SourenTestTool.spec (it bundles
REM  the pyvisa_py VISA backend). Do NOT run
REM  "pyinstaller ... gui_app.py" -- that overwrites the spec and
REM  the exe will fail with "cannot create ResourceManager".
REM ============================================================

cd /d "%~dp0"

echo [1/3] Checking pyvisa-py ...
python -c "import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('pyvisa_py') else 1)"
if errorlevel 1 (
    echo pyvisa_py not found, installing ...
    pip install pyvisa-py
)

echo [2/3] Checking pyqtgraph ...
python -c "import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('pyqtgraph') else 1)"
if errorlevel 1 (
    echo pyqtgraph not found, installing ...
    pip install pyqtgraph
)

echo [3/3] Building with spec (includes VISA backend + report charts) ...
pyinstaller --clean --noconfirm SourenTestTool.spec

echo.
echo Done. exe is at dist\SourenTestTool.exe
pause
