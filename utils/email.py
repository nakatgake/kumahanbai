import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from typing import List, Optional

from sqlalchemy.orm import Session
from database import SessionLocal
from models import SystemSetting


def _load_smtp_settings(db: Session) -> dict:
    """SystemSetting から SMTP 設定を取得し、辞書で返す。
    必要なキー: smtp_host, smtp_port, smtp_user, smtp_pass, smtp_from
    """
    keys = ["smtp_host", "smtp_port", "smtp_user", "smtp_pass", "smtp_from"]
    settings = {s.key: s.value for s in db.query(SystemSetting).filter(SystemSetting.key.in_(keys)).all()}
    # デフォルトポートは 587、送信元はユーザー名にフォールバック
    settings.setdefault("smtp_port", "587")
    settings.setdefault("smtp_from", settings.get("smtp_user", ""))
    return settings


def send_notification(
    subject: str,
    body: str,
    to: Optional[List[str]] = None,
    attachments: Optional[List[dict]] = None
) -> bool:
    """SystemSetting に登録された SMTP 設定でメールを送信。
    - `to` が None の場合は社内担当者 info@kumanomorikaken.co.jp に送信。
    - `attachments` は [{'name': str, 'content': bytes}] のリスト。
    - 例外が発生したら `False` を返す。
    """
    db = SessionLocal()
    try:
        s = _load_smtp_settings(db)
        host = s.get("smtp_host")
        port = int(s.get("smtp_port", 587))
        user = s.get("smtp_user")
        password = s.get("smtp_pass")
        sender = s.get("smtp_from") or user

        if not all([host, user, password]):
            raise RuntimeError("SMTP 設定が未完了です (host/user/pass が必要)")

        # Create message
        msg = MIMEMultipart()
        msg["Subject"] = subject
        msg["From"] = sender
        msg["To"] = ", ".join(to or ["info@kumanomorikaken.co.jp"])
        
        # Body
        msg.attach(MIMEText(body, 'plain'))

        # Attachments
        if attachments:
            for att in attachments:
                part = MIMEApplication(att['content'], Name=att['name'])
                part['Content-Disposition'] = f'attachment; filename="{att["name"]}"'
                msg.attach(part)

        context = ssl.create_default_context()
        if port == 465:
            server = smtplib.SMTP_SSL(host, port, timeout=10, context=context)
        else:
            server = smtplib.SMTP(host, port, timeout=10)
            server.starttls(context=context)
        server.login(user, password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"メール送信エラー: {e}")
        return False
    finally:
        db.close()


def send_admin_notification(subject: str, body: str) -> bool:
    """管理者通知メールを送信する。送信先はシステム設定の notification_email。
    send_admin_email_sync の代替として、セッション管理を内部で完結させる。
    """
    db = SessionLocal()
    try:
        notif_setting = db.query(SystemSetting).filter(SystemSetting.key == "notification_email").first()
        target = (notif_setting.value if notif_setting and notif_setting.value else None) or "info@kumanomorikaken.co.jp"
    finally:
        db.close()
    return send_notification(subject=subject, body=body, to=[target])
