@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [Git] 변경 사항 추가 중...
git add -A

echo [Git] 커밋 중...
git commit -m "자동 동기화: %date% %time%" 2>nul
if errorlevel 1 (
  echo 변경 사항이 없거나 이미 커밋되어 있습니다.
) else (
  echo [Git] 원격 저장소로 푸시 중...
  git push origin main
  if errorlevel 1 (
    echo 푸시 실패. 네트워크/권한을 확인하세요.
    pause
    exit /b 1
  )
  echo 완료.
)

pause
