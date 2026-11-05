"""Envoi des documents par e-mail (SMTP). Identifiants et destinataires : variables d'environnement / GitHub Secrets, jamais dans le code."""
import os
import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path

XLSX = ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
DOCX = ("application", "vnd.openxmlformats-officedocument.wordprocessingml.document")


class MailError(RuntimeError):
    pass


def recipients(cfg) -> list[str]:
    raw = os.environ.get("MAIL_TO") or ",".join(cfg.data.get("email", {}).get("destinataires", []))
    return [a.strip() for a in raw.split(",") if a.strip()]


def build_message(subject: str, body: str, attachments: list[Path], sender: str, to: list[str]) -> EmailMessage:
    msg = EmailMessage()
    msg["Subject"], msg["From"], msg["To"] = subject, sender, ", ".join(to)
    msg.set_content(body)
    for f in attachments:
        maintype, subtype = XLSX if f.suffix == ".xlsx" else DOCX if f.suffix == ".docx" else ("application", "octet-stream")
        msg.add_attachment(Path(f).read_bytes(), maintype=maintype, subtype=subtype, filename=Path(f).name)
    return msg


def send(msg: EmailMessage) -> None:
    host, user, password = os.environ.get("SMTP_HOST"), os.environ.get("SMTP_USER"), os.environ.get("SMTP_PASSWORD")
    if not (host and user and password):
        raise MailError("Configuration SMTP incomplète : SMTP_HOST, SMTP_USER et SMTP_PASSWORD sont requis.")
    port = int(os.environ.get("SMTP_PORT", "587"))
    ctx = ssl.create_default_context()
    try:
        if port == 465:
            with smtplib.SMTP_SSL(host, port, context=ctx, timeout=60) as s:
                s.login(user, password), s.send_message(msg)
        else:
            with smtplib.SMTP(host, port, timeout=60) as s:
                s.starttls(context=ctx), s.login(user, password), s.send_message(msg)
    except (smtplib.SMTPException, OSError) as e:
        raise MailError(f"Envoi impossible : {e}") from e
