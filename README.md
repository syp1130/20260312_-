# 도미노피자 재고·발주 자동화

재고 엑셀을 분석해 **재고 부족**인 품목을 찾고, 담당 거래처에 **발주 메일**을 보내는 웹 시스템입니다.

## 기능

- **웹 재고 입력**: 기준 데이터 불러오기 후 현재재고만 입력해 분석
- **엑셀 업로드**: 엑셀 파일로 재고 분석
- **재고 분석**: 현재재고 < 안전재고 → 발주 필요, 권장 수량 = MAX(MOQ, 안전재고 - 현재재고)
- **발주 메일 발송**: 거래처별로 발주서를 Gmail로 발송
- **팀 비밀번호**: 배포 시 `TEAM_PASSWORD` 환경변수로 접속 제한

## 로컬 실행

1. 의존성 설치:
   ```bash
   pip install -r requirements.txt
   ```

2. (선택) Gmail 앱 비밀번호: `.env`에 `GMAIL_APP_PASSWORD=...` 또는 환경변수 설정

3. (선택) 팀 비밀번호: 환경변수 `TEAM_PASSWORD` 설정 시 로그인 필요. 비우면 비밀번호 없이 접속.

4. 서버 실행:
   ```bash
   python app.py
   ```
   브라우저에서 http://localhost:5000 접속

## Vercel 배포

1. [Vercel](https://vercel.com)에 로그인 후 **New Project** → GitHub 저장소 `syp1130/20260312_-` 연결

2. **Environment Variables**에 설정:
   - `TEAM_PASSWORD`: 팀원만 알 수 있는 접속 비밀번호 (필수 권장)
   - `SESSION_SECRET`: 세션 암호화용 랜덤 문자열 (예: `openssl rand -hex 32`)
   - `GMAIL_APP_PASSWORD`: Gmail 앱 비밀번호 (발주 메일 발송 시 필요)

3. **Deploy** 후 배포 URL로 접속 → 비밀번호 입력 후 사용

## 저장소

- GitHub: https://github.com/syp1130/20260312_-.git

## 설정

- 발송 메일: `config.py`의 `SENDER_EMAIL`
- 점포명: `config.py`의 `STORE_NAME`
