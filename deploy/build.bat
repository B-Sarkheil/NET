@echo off
cd /d "%~dp0"
echo =============================================
echo  Namvaran Excel Tracker - Build EXE
echo =============================================
echo.

echo [1/4] Installing app dependencies...
python -m pip install --quiet -r "..\src\requirements.txt"
if errorlevel 1 (
    echo ERROR: Python not found, or failed to install dependencies from src\requirements.txt.
    pause & exit /b 1
)

echo [2/4] Installing PyInstaller...
python -m pip install --quiet pyinstaller
if errorlevel 1 (
    echo ERROR: Failed to install PyInstaller.
    pause & exit /b 1
)

echo [3/4] Building NET.exe...
python -m PyInstaller --onedir --clean --noconfirm ^
    --add-data "..\src\templates;templates" ^
    --add-data "..\src\static;static" ^
    --collect-submodules openpyxl ^
    --distpath "..\product" ^
    --workpath ".\build_work" ^
    --specpath "." ^
    --name NET ^
    ..\src\app.py

if errorlevel 1 (
    echo ERROR: Build failed.
    pause & exit /b 1
)

echo [4/4] Copying config.txt to product...
copy /Y "..\src\config.txt" "..\product\NET\config.txt" > nul

echo.
echo =============================================
echo  Build successful!
echo  Output: product\NET\
echo  1. Copy entire product\NET\ folder to target computer
echo  2. Edit config.txt with correct server paths
echo  3. Run NET.exe
echo =============================================
pause
