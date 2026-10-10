@echo off
cd /d "%~dp0"
title GitHub Push - Gangbuk SmartFarm
echo ====================================================
echo  [SmartFarm] GitHub 저장소로 코드 전송 중...
echo  화면에 GitHub 로그인 창이 뜨면 로그인해 주세요!
echo ====================================================
echo.
git branch -M main
git push -u origin main
echo.
if %ERRORLEVEL% EQU 0 (
    echo ====================================================
    echo  [성공] GitHub로 코드가 성공적으로 업로드되었습니다!
    echo ====================================================
) else (
    echo [안내] 전송이 완료되지 않았습니다. 창을 닫고 다시 시도해 주세요.
)
echo.
pause
