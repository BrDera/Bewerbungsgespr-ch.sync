"""
Bewerbungs-Mail-Sync (Cloud-Version)

Durchsucht per IMAP den Sent-Ordner einer iCloud-Adresse nach Mails der
letzten 2 Tage, filtert nach Stichwoertern im Betreff und traegt Treffer
als Karte (Status "Offen") in die Supabase-Tabelle "bewerbungen" ein.

Gedacht zum Ausfuehren in GitHub Actions nach Zeitplan - siehe
mail-sync.yml. Benoetigte Umgebungsvariablen (als GitHub Secrets):
  ICLOUD_EMAIL         volle iCloud-Adresse, z.B. name@icloud.com
  ICLOUD_APP_PASSWORD  App-spezifisches Passwort von appleid.apple.com
  SUPABASE_URL         z.B. https://xxxx.supabase.co
  SUPABASE_ANON_KEY    der "anon public" Key aus Supabase

Nutzt nur die Python-Standardbibliothek, kein pip install noetig.
"""

import imaplib
import email
import re
import os
import sys
import json
import hashlib
import datetime
import urllib.request
from email.header import decode_header
from email.utils import parseaddr, parsedate_to_datetime

IMAP_HOST = "imap.mail.me.com"
IMAP_PORT = 993

# Stichwoerter, die im Betreff auf eine Bewerbung hindeuten (klein geschrieben).
# Bei Bedarf ergaenzen/anpassen.
KEYWORDS = ["bewerbung", "anschreiben", "application", "praktikum", "initiativbewerbung"]

CANDIDATE_SENT_FOLDERS = ["Sent Messages", "Sent", "INBOX.Sent Messages"]


def decode_str(raw):
    if raw is None:
        return ""
    parts = decode_header(raw)
    out = ""
    for text, enc in parts:
        if isinstance(text, bytes):
            out += text.decode(enc or "utf-8", errors="replace")
        else:
            out += text
    return out


def connect():
    email_addr = os.environ["ICLOUD_EMAIL"]
    password = os.environ["ICLOUD_APP_PASSWORD"]
    conn = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    conn.login(email_addr, password)
    return conn


def list_folder_names(conn):
    typ, mailboxes = conn.list()
    names = []
    if typ == "OK":
        for m in mailboxes:
            decoded = m.decode(errors="replace")
            match = re.search(r'"([^"]+)"$', decoded)
            if match:
                names.append(match.group(1))
    return names


def select_sent_folder(conn):
    names = list_folder_names(conn)

    for candidate in CANDIDATE_SENT_FOLDERS:
        if candidate in names:
            typ, _ = conn.select(f'"{candidate}"', readonly=True)
            if typ == "OK":
                return candidate

    for name in names:
        if "sent" in name.lower():
            typ, _ = conn.select(f'"{name}"', readonly=True)
            if typ == "OK":
                return name

    raise RuntimeError("Kein Sent-Ordner gefunden. Gefundene Ordner: " + ", ".join(names))


def fetch_recent_messages(conn, days=2):
    since_date = datetime.date.today() - datetime.timedelta(days=days)
    date_str = since_date.strftime("%d-%b-%Y")
    typ, data = conn.search(None, f"(SINCE {date_str})")
    if typ != "OK" or not data or data[0] is None:
        return []
    ids = data[0].split()
    messages = []
    for msg_id in ids:
        typ, msg_data = conn.fetch(msg_id, "(RFC822)")
        if typ != "OK" or not msg_data or msg_data[0] is None:
            continue
        raw = msg_data[0][1]
        messages.append(email.message_from_bytes(raw))
    return messages


def matches_keywords(subject):
    lower = subject.lower()
    return any(kw in lower for kw in KEYWORDS)


def upsert_card(supabase_url, supabase_key, message_id, subject, recipient, iso_date):
    card_id = "mail-" + hashlib.sha1(message_id.encode("utf-8")).hexdigest()[:16]
    payload = {
        "id": card_id,
        "data": {
            "org": recipient or "Unbekannt",
            "pos": subject,
            "bereich": "Sonstiges",
            "person": "",
            "mail": recipient or "",
            "link": "",
            "datum": iso_date,
            "status": "yellow",
            "notes": "Automatisch erkannt aus Sent-Ordner (IMAP)",
        },
    }
    endpoint = supabase_url.rstrip("/") + "/rest/v1/bewerbungen?on_conflict=id"
    req = urllib.request.Request(
        endpoint,
        data=json.dumps(payload).encode("utf-8"),
        method="POST",
        headers={
            "apikey": supabase_key,
            "Authorization": f"Bearer {supabase_key}",
            "Content-Type": "application/json",
            "Prefer": "resolution=ignore-duplicates,return=minimal",
        },
    )
    try:
        urllib.request.urlopen(req, timeout=15)
    except Exception as e:
        print(f"Fehler beim Speichern von {card_id}: {e}", file=sys.stderr)


def main():
    supabase_url = os.environ["SUPABASE_URL"]
    supabase_key = os.environ["SUPABASE_ANON_KEY"]

    conn = connect()
    try:
        folder = select_sent_folder(conn)
        print(f"Nutze Ordner: {folder}")

        messages = fetch_recent_messages(conn, days=2)
        print(f"{len(messages)} Nachricht(en) der letzten 2 Tage gefunden.")

        added = 0
        for msg in messages:
            subject = decode_str(msg.get("Subject"))
            if not matches_keywords(subject):
                continue

            message_id = msg.get("Message-ID") or f"{msg.get('Date')}-{subject}"
            to_header = msg.get("To", "")
            _, addr = parseaddr(to_header)
            recipient = addr or to_header

            try:
                dt = parsedate_to_datetime(msg.get("Date"))
                iso_date = dt.date().isoformat()
            except Exception:
                iso_date = ""

            upsert_card(supabase_url, supabase_key, message_id, subject, recipient, iso_date)
            added += 1

        print(f"{added} passende Bewerbungs-Mail(s) verarbeitet.")
    finally:
        conn.logout()


if __name__ == "__main__":
    main()
