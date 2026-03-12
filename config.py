import os

# 발주 메일 발송용 계정 (Gmail)
# 비밀번호는 반드시 환경변수 GMAIL_APP_PASSWORD 로 설정하세요.
SENDER_EMAIL = "sypark.dpk@gmail.com"
# Gmail 앱 비밀번호: https://myaccount.google.com/apppasswords 에서 발급

# 기본 엑셀 파일
DEFAULT_EXCEL_PATH = "domino_inventory_training.xlsx"

# 점포명 (이메일 템플릿용)
STORE_NAME = "도미노피자 점포"

# 팀 접속 비밀번호 (Vercel 등에서는 환경변수 TEAM_PASSWORD 로 설정)
# 비어 있으면 비밀번호 없이 접속 가능 (로컬용)
TEAM_PASSWORD = os.environ.get("TEAM_PASSWORD", "").strip()
SESSION_SECRET = os.environ.get("SESSION_SECRET", "inventory-order-secret-change-in-production")
