import imaplib
import email
import base64
import json
import argparse
import os
import sys
from email.header import decode_header
from datetime import datetime
from typing import Optional

import requests


AUTHORITY = "https://login.microsoftonline.com/common"
TOKEN_ENDPOINT = f"{AUTHORITY}/oauth2/v2.0/token"
IMAP_HOST = "outlook.office365.com"
IMAP_PORT = 993
SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All offline_access"
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


def load_config() -> dict:
    """Load settings from config.json next to the script."""
    if not os.path.exists(CONFIG_PATH):
        return {}
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def get_access_token(client_id: str, refresh_token: str) -> dict:
    """Exchange refresh_token for a new access_token."""
    data = {
        "client_id": client_id,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "scope": SCOPE,
    }
    resp = requests.post(TOKEN_ENDPOINT, data=data, timeout=30)
    if resp.status_code != 200:
        print(f"Token error: {resp.status_code} {resp.text}", file=sys.stderr)
        sys.exit(1)
    return resp.json()


def build_xoauth2_string(user: str, access_token: str) -> str:
    """Build XOAUTH2 authentication string."""
    auth_string = f"user={user}\x01auth=Bearer {access_token}\x01\x01"
    return auth_string


def decode_mime_header(value: str) -> str:
    """Decode a MIME-encoded header value."""
    if value is None:
        return ""
    parts = decode_header(value)
    decoded = []
    for part, charset in parts:
        if isinstance(part, bytes):
            decoded.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            decoded.append(part)
    return "".join(decoded)


def get_text_body(msg: email.message.Message) -> str:
    """Extract plain text body from an email message."""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            disposition = str(part.get("Content-Disposition", ""))
            if content_type == "text/plain" and "attachment" not in disposition:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    return payload.decode(charset, errors="replace")
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            return payload.decode(charset, errors="replace")
    return ""


def get_html_body(msg: email.message.Message) -> str:
    """Extract HTML body and strip tags to plain text as fallback."""
    import re
    html = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/html" and "attachment" not in str(part.get("Content-Disposition", "")):
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    html = payload.decode(charset, errors="replace")
                    break
    elif msg.get_content_type() == "text/html":
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            html = payload.decode(charset, errors="replace")

    if not html:
        return ""

    # Strip HTML tags to get readable text
    html = re.sub(r'<style[^>]*>.*?</style>', '', html, flags=re.DOTALL)
    html = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.DOTALL)
    html = re.sub(r'<br\s*/?>', '\n', html, flags=re.IGNORECASE)
    html = re.sub(r'</p>', '\n\n', html, flags=re.IGNORECASE)
    html = re.sub(r'</div>', '\n', html, flags=re.IGNORECASE)
    html = re.sub(r'</tr>', '\n', html, flags=re.IGNORECASE)
    html = re.sub(r'<[^>]+>', '', html)
    html = re.sub(r'&nbsp;', ' ', html)
    html = re.sub(r'&amp;', '&', html)
    html = re.sub(r'&lt;', '<', html)
    html = re.sub(r'&gt;', '>', html)
    html = re.sub(r'&#\d+;', '', html)
    html = re.sub(r'\n{3,}', '\n\n', html)
    return html.strip()


def get_attachments(msg: email.message.Message) -> list:
    """Return list of attachment filenames and sizes."""
    attachments = []
    if not msg.is_multipart():
        return attachments
    for part in msg.walk():
        disposition = str(part.get("Content-Disposition", ""))
        if "attachment" in disposition or "inline" in disposition:
            filename = part.get_filename()
            if filename:
                filename = decode_mime_header(filename)
                payload = part.get_payload(decode=True)
                size = len(payload) if payload else 0
                attachments.append((filename, size))
    return attachments


def format_size(size_bytes: int) -> str:
    """Format bytes to human readable size."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"


def format_message(i: int, total: int, msg: email.message.Message) -> str:
    """Format a single email message as readable text."""
    subject = decode_mime_header(msg["Subject"])
    from_addr = decode_mime_header(msg["From"])
    to_addr = decode_mime_header(msg["To"])
    cc_addr = decode_mime_header(msg.get("Cc", ""))
    date_str = msg["Date"] or ""

    lines = []
    lines.append("=" * 70)
    lines.append(f"  Письмо {i} из {total}")
    lines.append("=" * 70)
    lines.append(f"  Тема:       {subject}")
    lines.append(f"  От:         {from_addr}")
    lines.append(f"  Кому:       {to_addr}")
    if cc_addr:
        lines.append(f"  Копия:      {cc_addr}")
    lines.append(f"  Дата:       {date_str}")

    attachments = get_attachments(msg)
    if attachments:
        att_list = ", ".join(f"{name} ({format_size(size)})" for name, size in attachments)
        lines.append(f"  Вложения:   {att_list}")

    lines.append("-" * 70)

    body = get_text_body(msg)
    if not body.strip():
        body = get_html_body(msg)

    if body.strip():
        lines.append(body.strip())
    else:
        lines.append("  (нет текстового содержимого)")

    lines.append("")
    return "\n".join(lines)


def fetch_all_messages(
    client_id: str,
    refresh_token: str,
    user_email: Optional[str] = None,
) -> tuple:
    """Fetch all messages from Outlook inbox via IMAP with OAuth2.
    Returns (user_email, list_of_formatted_messages, total_count).
    """
    # 1. Get access token
    print("Getting access token...")
    token_data = get_access_token(client_id, refresh_token)
    access_token = token_data["access_token"]

    # Try to extract email from id_token or access_token (both are JWTs)
    if not user_email:
        for token_key in ("id_token", "access_token"):
            token_val = token_data.get(token_key)
            if not token_val:
                continue
            try:
                payload_b64 = token_val.split(".")[1]
                payload_b64 += "=" * (4 - len(payload_b64) % 4)
                claims = json.loads(base64.urlsafe_b64decode(payload_b64))
                user_email = (
                    claims.get("preferred_username")
                    or claims.get("email")
                    or claims.get("upn")
                    or claims.get("unique_name")
                )
                if user_email:
                    break
            except Exception:
                continue

    if not user_email:
        print("Cannot determine user email. Set 'email' in config.", file=sys.stderr)
        return None, [], 0

    print(f"Authenticated as: {user_email}")

    # 2. Connect to IMAP
    print("Connecting to IMAP...")
    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)

    auth_string = build_xoauth2_string(user_email, access_token)
    imap.authenticate("XOAUTH2", lambda _: auth_string.encode())
    print("IMAP authentication successful.")

    # 3. Select INBOX
    imap.select("INBOX")
    status, data = imap.search(None, "ALL")
    if status != "OK":
        print("Failed to search messages.", file=sys.stderr)
        imap.logout()
        return user_email, [], 0

    msg_ids = data[0].split()
    print(f"Total messages in INBOX: {len(msg_ids)}")

    # 4. Fetch each message
    all_formatted = []
    for i, msg_id in enumerate(msg_ids, 1):
        status, msg_data = imap.fetch(msg_id, "(RFC822)")
        if status != "OK":
            print(f"  [{i}] Failed to fetch message {msg_id.decode()}")
            continue

        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        subject = decode_mime_header(msg["Subject"])
        from_addr = decode_mime_header(msg["From"])
        date_str = msg["Date"] or ""

        formatted = format_message(i, len(msg_ids), msg)
        all_formatted.append(formatted)

        print(f"  [{i}/{len(msg_ids)}] {date_str[:16]}  {from_addr[:30]}  {subject[:40]}")

    # 5. Logout
    imap.logout()

    return user_email, all_formatted, len(msg_ids)


def write_report(user_email: str, formatted: list, total: int, output_dir: str):
    """Write formatted messages to {email}.txt."""
    os.makedirs(output_dir, exist_ok=True)
    report_path = os.path.join(output_dir, f"{user_email}.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(f"Почтовый ящик: {user_email}\n")
        f.write(f"Дата выгрузки:  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Всего писем:    {total}\n\n")
        f.write("\n".join(formatted))
    print(f"Сохранено: {report_path}")


def main():
    config = load_config()

    parser = argparse.ArgumentParser(
        description="Fetch all Outlook emails via IMAP using OAuth2 refresh token. "
        "Supports multiple accounts via config.json."
    )
    parser.add_argument("--client-id", default=None, help="Azure app client_id (single account mode)")
    parser.add_argument("--refresh-token", default=None, help="OAuth2 refresh_token (single account mode)")
    parser.add_argument("--email", default=None, help="User email (auto-detected if not provided)")
    parser.add_argument("--output-dir", default=None, help="Directory to save .txt files (if not set, prints to console)")
    args = parser.parse_args()

    output_dir = args.output_dir or config.get("output_dir")

    # Build list of accounts to process
    accounts = []
    if args.client_id and args.refresh_token:
        accounts.append({
            "client_id": args.client_id,
            "refresh_token": args.refresh_token,
            "email": args.email,
        })
    elif config.get("accounts"):
        accounts = config["accounts"]
    elif config.get("client_id") and config.get("refresh_token"):
        # Backward compat: old single-account config format
        accounts.append({
            "client_id": config["client_id"],
            "refresh_token": config["refresh_token"],
            "email": config.get("email"),
        })

    if not accounts:
        print("Error: no accounts configured. Add 'accounts' to config.json or pass --client-id/--refresh-token.", file=sys.stderr)
        sys.exit(1)

    print(f"Accounts to process: {len(accounts)}\n")

    for idx, account in enumerate(accounts, 1):
        client_id = account.get("client_id")
        refresh_token = account.get("refresh_token")
        email_addr = account.get("email")

        if not client_id or not refresh_token:
            print(f"[Account {idx}] Skipping: missing client_id or refresh_token", file=sys.stderr)
            continue

        print(f"{'=' * 50}")
        print(f"  Account {idx}/{len(accounts)}: {email_addr or '(auto-detect)'}")
        print(f"{'=' * 50}")

        try:
            user_email, formatted, total = fetch_all_messages(client_id, refresh_token, email_addr)
        except Exception as e:
            print(f"[Account {idx}] Error: {e}", file=sys.stderr)
            continue

        if not user_email:
            continue

        if output_dir and formatted:
            write_report(user_email, formatted, total, output_dir)
        elif formatted:
            print()
            print("\n".join(formatted))

        print(f"Готово. Обработано писем: {total}.\n")


if __name__ == "__main__":
    main()
