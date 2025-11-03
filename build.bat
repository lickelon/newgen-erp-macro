@echo off
echo ====================================
echo Building Executable
echo ====================================
echo.

echo [1/5] Checking for running processes...
tasklist | find /i "newgen_erp_macro.exe" >nul
if not errorlevel 1 (
    echo Stopping running newgen_erp_macro.exe...
    taskkill /f /im newgen_erp_macro.exe >nul 2>&1
    echo Waiting for process to fully terminate...
    timeout /t 5 /nobreak >nul
    echo Done.
) else (
    echo No running process found.
)
echo.

echo [2/5] Cleaning previous build...
if exist "dist\newgen_erp_macro.exe" (
    echo Removing old executable...

    REM Try normal delete first
    del /f /q "dist\newgen_erp_macro.exe" 2>nul
    timeout /t 1 /nobreak >nul

    REM If still exists, try PowerShell force delete
    if exist "dist\newgen_erp_macro.exe" (
        echo Retrying with PowerShell...
        powershell -Command "Remove-Item -Path 'dist\newgen_erp_macro.exe' -Force -ErrorAction SilentlyContinue" 2>nul
        timeout /t 1 /nobreak >nul
    )

    REM Final check
    if exist "dist\newgen_erp_macro.exe" (
        echo Warning: Could not delete old executable.
        echo Please close any programs using the file and try again.
        pause
        exit /b 1
    )
)
if exist "dist" rmdir /s /q "dist" 2>nul
if exist "build" rmdir /s /q "build" 2>nul
echo Done.

echo.
echo [3/5] Running PyInstaller...
uv run pyinstaller gui_app.spec --clean
if errorlevel 1 (
    echo Build failed!
    pause
    exit /b 1
)
echo Done.

echo.
echo [4/5] Verifying executable...
if exist "dist\newgen_erp_macro.exe" (
    echo Success: dist\newgen_erp_macro.exe
) else (
    echo Failed to create executable!
    pause
    exit /b 1
)

echo.
echo [5/5] Cleaning temporary files...
if exist "build" rmdir /s /q "build"
echo Done.

echo.
echo ====================================
echo Build Complete!
echo ====================================
echo Executable: dist\newgen_erp_macro.exe
echo.
pause
