# 도미노피자 재고·발주 자동화

재고 엑셀을 분석해 **재고 부족**인 품목을 찾고, 담당 거래처에 **발주 메일**을 보내는 웹 시스템입니다.

## 기능

- **데이터 입력**: 엑셀 파일 업로드 (또는 기본 `domino_inventory_training.xlsx` 사용)
- **재고 분석**: 현재재고 < 안전재고 → 발주 필요, 권장 수량 = MAX(MOQ, 안전재고 - 현재재고)
- **발주 메일 발송**: 거래처별로 발주서 내용을 **sypark.dpk@gmail.com** 계정으로 발송

## 실행 방법

1. 가상환경 권장 후 의존성 설치:
   ```bash
   pip install -r requirements.txt
   ```

2. **Gmail 앱 비밀번호** 설정 (Gmail 2단계 인증 후 [앱 비밀번호](https://myaccount.google.com/apppasswords) 발급):
   - Windows CMD: `set GMAIL_APP_PASSWORD=발급받은16자리비밀번호`
   - PowerShell: `$env:GMAIL_APP_PASSWORD="발급받은16자리비밀번호"`

3. 서버 실행:
   ```bash
   python app.py
   ```

4. 브라우저에서 **http://localhost:5000** 접속 후:
   - 엑셀 업로드(선택) → **재고 분석 실행** → 결과 확인 → **발주 메일 발송**

## 엑셀 구조 (참고)

- **Inventory**: 품목코드, 재료명, 규격, 단위, 현재재고, 안전재고, MOQ, 거래처, 거래처이메일 등
- **Suppliers**: 거래처명, 담당자, 이메일, 리드타임
- **EmailTemplate**: 발주 메일 제목/본문 템플릿

## 설정

- 발송 메일 계정: `config.py`의 `SENDER_EMAIL` (기본: sypark.dpk@gmail.com)
- 점포명: `config.py`의 `STORE_NAME`
