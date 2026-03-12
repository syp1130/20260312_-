@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo 재고 발주 웹 서버를 시작합니다...
echo 브라우저가 곧 열립니다. (안 열리면 http://localhost:5000 입력)
echo.
echo 종료하려면 이 창을 닫거나 아무 키나 누르세요.
echo.

start "" cmd /c "ping -n 3 127.0.0.1 > nul && start http://localhost:5000"
python app.py

pause
