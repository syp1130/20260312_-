@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [Git] 원격 저장소에서 최신 내용 가져오는 중...
git pull origin main --no-edit 2>nul
if errorlevel 1 (
  echo pull 생략 또는 실패. 로컬 변경 우선 진행.
)

echo.
echo [Git] 변경 사항 추가 중...
git add -A

echo [Git] 커밋 중...
git commit -m "자동 동기화: %date% %time%" 2>nul
if errorlevel 1 (
  echo 변경 사항이 없거나 이미 커밋되어 있습니다.
  goto PUSH
)

:PUSH
echo [Git] GitHub로 푸시 중...
git push origin main
if errorlevel 1 (
  echo.
  echo 푸시 실패. 아래를 확인하세요.
  echo - 인터넷 연결
  echo - GitHub 로그인 (또는 자격 증명)
  echo - 저장소 주소: https://github.com/syp1130/20260312_1
  pause
  exit /b 1
)

echo.
echo === GitHub 반영 완료 ===
pause
