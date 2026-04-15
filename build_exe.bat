@echo off
:: ============================================================
:: HWP 뷰어  -  Windows 단일 실행파일(.exe) 빌드 스크립트
:: 인터넷 가능한 PC에서 실행 후 dist\HWP뷰어.exe 를 내부망에 복사
:: ============================================================

echo [1/3] 의존성 설치 중...
pip install -r requirements.txt
if errorlevel 1 (
    echo 의존성 설치 실패
    pause & exit /b 1
)

echo.
echo [2/3] pyhwp 데이터 파일 경로 확인 중...
for /f "tokens=*" %%i in ('python -c "import hwp5, os; print(os.path.dirname(hwp5.__file__))"') do set HWP5_DIR=%%i
echo pyhwp 경로: %HWP5_DIR%

echo.
echo [3/3] 단일 exe 빌드 중...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "HWP뷰어" ^
    --add-data "%HWP5_DIR%;hwp5" ^
    --hidden-import tkinterweb ^
    --hidden-import olefile ^
    --hidden-import hwp5 ^
    --hidden-import hwp5.hwp5html ^
    --hidden-import lxml ^
    --hidden-import lxml.etree ^
    hwp_viewer.py

echo.
if exist dist\HWP뷰어.exe (
    echo ====================================
    echo  빌드 성공!
    echo  실행파일: dist\HWP뷰어.exe
    echo  이 파일 하나만 내부망에 복사하세요.
    echo ====================================
) else (
    echo 빌드 실패. 위 오류 메시지를 확인하세요.
)
pause
