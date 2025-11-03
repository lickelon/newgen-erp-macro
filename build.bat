@echo off
chcp 65001 > nul
echo ====================================
echo 부양가족 대량입력 실행파일 빌드
echo ====================================
echo.

echo [1/4] 이전 빌드 정리 중...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
echo ✓ 정리 완료

echo.
echo [2/4] PyInstaller 실행 중...
uv run pyinstaller gui_app.spec --clean
if errorlevel 1 (
    echo ✗ 빌드 실패
    pause
    exit /b 1
)
echo ✓ 빌드 완료

echo.
echo [3/4] 실행파일 확인 중...
if exist "dist\부양가족_대량입력.exe" (
    echo ✓ 실행파일 생성 성공: dist\부양가족_대량입력.exe
) else (
    echo ✗ 실행파일 생성 실패
    pause
    exit /b 1
)

echo.
echo [4/4] 빌드 임시 파일 정리 중...
if exist "build" rmdir /s /q "build"
echo ✓ 정리 완료

echo.
echo ====================================
echo 빌드 완료! 🎉
echo ====================================
echo 실행파일 위치: dist\부양가족_대량입력.exe
echo.
pause
