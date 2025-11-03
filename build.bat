@echo off
echo ====================================
echo Building Executable
echo ====================================
echo.

echo [1/4] Cleaning previous build...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
echo Done.

echo.
echo [2/4] Running PyInstaller...
uv run pyinstaller gui_app.spec --clean
if errorlevel 1 (
    echo Build failed!
    pause
    exit /b 1
)
echo Done.

echo.
echo [3/4] Verifying executable...
if exist "dist\BuYangGaJok_Input.exe" (
    echo Success: dist\BuYangGaJok_Input.exe
) else (
    echo Failed to create executable!
    pause
    exit /b 1
)

echo.
echo [4/4] Cleaning temporary files...
if exist "build" rmdir /s /q "build"
echo Done.

echo.
echo ====================================
echo Build Complete!
echo ====================================
echo Executable: dist\BuYangGaJok_Input.exe
echo.
pause
