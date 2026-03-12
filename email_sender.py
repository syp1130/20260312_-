# -*- coding: utf-8 -*-
"""발주 이메일 발송 (Gmail SMTP)."""
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import os

from config import SENDER_EMAIL, STORE_NAME


def get_gmail_password():
    """환경변수에서 Gmail 앱 비밀번호 읽기."""
    return os.environ.get("GMAIL_APP_PASSWORD", "").strip()


def fill_template(template_body: str, template_subject: str, supplier_name: str, item_list_html: str) -> tuple[str, str]:
    """템플릿에 변수 치환."""
    order_date = datetime.now().strftime("%Y-%m-%d")
    body = (
        template_body.replace("{{SUPPLIER_NAME}}", supplier_name)
        .replace("{{STORE_NAME}}", STORE_NAME)
        .replace("{{ORDER_DATE}}", order_date)
        .replace("{{ITEM_LIST}}", item_list_html)
        .replace("{{INTERNAL_OWNER}}", "담당자")
        .replace("\\n", "\n")
    )
    subject = (
        template_subject.replace("{{STORE_NAME}}", STORE_NAME)
        .replace("{{SUPPLIER_NAME}}", supplier_name)
        .replace("{{ORDER_DATE}}", order_date)
    )
    return subject, body


def send_order_email(
    to_email: str,
    supplier_name: str,
    item_list: list[dict],
    template_subject: str,
    template_body: str,
) -> tuple[bool, str]:
    """
    발주 이메일 발송.
    item_list: [{"재료명": "", "규격": "", "단위": "", "발주권장수량": 0}, ...]
    """
    # 품목 목록 HTML
    lines = []
    for item in item_list:
        name = item.get("재료명", "")
        spec = item.get("규격", "")
        unit = item.get("단위", "개")
        qty = item.get("발주권장수량", 0)
        lines.append(f"- {name} ({spec}) : {qty}{unit}")
    item_list_html = "<br>".join(lines) if lines else "(없음)"

    subject, body = fill_template(template_body, template_subject, supplier_name, item_list_html)
    body_html = body.replace("\n", "<br>")

    password = get_gmail_password()
    if not password:
        return False, "GMAIL_APP_PASSWORD 환경변수가 설정되지 않았습니다."

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = SENDER_EMAIL
    msg["To"] = to_email
    msg.attach(MIMEText(body, "plain", "utf-8"))
    msg.attach(MIMEText(body_html, "html", "utf-8"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(SENDER_EMAIL, password)
            server.sendmail(SENDER_EMAIL, [to_email], msg.as_string())
        return True, "발송 완료"
    except Exception as e:
        return False, str(e)
