"""
Bewerbungs-Mail-Sync (Cloud-Version, regelbasiert - keine bezahlte KI-API)

Durchsucht per IMAP den Sent-Ordner einer iCloud-Adresse nach Mails der
letzten 2 Tage, filtert nach Stichwoertern im Betreff und im Text,
gruppiert Treffer PRO FIRMA (per Empfaenger-Domain) und pflegt dafuer
jeweils genau eine Zeile in der Supabase-Tabelle "bewerbungen" inkl.
Mailverlauf. Ausserdem wird pro Empfaenger ein Eintrag in der Tabelle
"kontakte" gepflegt (fuer den Netzwerk-Tab im Board).

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

# Stichwoerter, die im Betreff ODER im Text auf eine Bewerbung hindeuten
# (klein geschrieben). Bei Bedarf ergaenzen/anpassen.
KEYWORDS = ["bewerbung", "anschreiben", "application", "praktikum", "initiativbewerbung"]

CANDIDATE_SENT_FOLDERS = ["Sent Messages", "Sent", "INBOX.Sent Messages"]

MAX_BODY_CHARS = 20000

CATEGORIES = [
    "Journalismus",
    "Gruendung/Gruendungszentrum",
    "Wissenschaft/Akademisch",
    "Start-up",
    "KI",
    "Kurzfristig/Nebenjob",
    "Sonstiges",
]

# Regelbasierte Kategorie-Zuordnung: Schluesselwort -> Kategorie.
# Wird gegen Betreff + Text + Empfaenger-Domain geprueft, erster Treffer gewinnt.
CATEGORY_RULES = [
    ("Journalismus", ["redaktion", "journalismus", "journalist", "presse", "rundfunk", "podcast", "nachrichten"]),
    ("Start-up", ["start-up", "startup", "founder", "gruender", "cfo", "coo", "cto"]),
    ("Gruendung/Gruendungszentrum", ["stiftung", "accelerator", "inkubator", "gruendungszentrum", "entrepreneurship"]),
    ("Wissenschaft/Akademisch", ["universitaet", "uni ", "hochschule", "forschung", "wissenschaftlich", "lehrstuhl", "institut"]),
    ("KI", ["kuenstliche intelligenz", " ki ", "artificial intelligence", "machine learning", " ai "]),
    ("Kurzfristig/Nebenjob", ["nebenjob", "aushilfe", "kellner", "kellnerin", "lieferant", "fahrer", "cafe", "café"]),
]

# Wörter, die auf eine Positions-/Rollenbezeichnung hindeuten - werden fuer
# den Untertitel aus Betreff/Text herausgesucht.
POSITION_HINTS = [
    "praktikum", "praktikant", "werkstudent", "volontariat", "trainee",
    "associate", "projekt-manager", "projektmanager", "programm-manager",
    "kursleiter", "redaktion", "redakteur",
]


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


def extract_body_text(msg):
    """Holt den Klartext-Inhalt einer Mail (auch bei mehrteiligen Mails)."""
    text_parts = []
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))
            if content_type == "text/plain" and "attachment" not in content_disposition.lower():
                try:
                    payload = part.get_payload(decode=True)
                    if not payload:
                        continue
                    charset = part.get_content_charset() or "utf-8"
                    text_parts.append(payload.decode(charset, errors="replace"))
                except Exception:
                    continue
    else:
        try:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or "utf-8"
                text_parts.append(payload.decode(charset, errors="replace"))
        except Exception:
            pass

    combined = "\n".join(text_parts)
    return combined[:MAX_BODY_CHARS]


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


def matches_keywords(text):
    lower = text.lower()
    return any(kw in lower for kw in KEYWORDS)


def clean_headline(subject):
    """Entfernt Re:/AW:/Fwd: und generische Praefixe aus dem Betreff."""
    cleaned = subject.strip()
    cleaned = re.sub(r"^(re|aw|fwd|wg)\s*:\s*", "", cleaned, flags=re.IGNORECASE).strip()
    # Haeufiges Muster: "Your application at X - <eigentliches Thema>"
    match = re.search(r"[-\u2013]\s*(.+)$", cleaned)
    if match and len(match.group(1)) >= 4:
        cleaned = match.group(1).strip()
    if len(cleaned) > 60:
        cleaned = cleaned[:57].rstrip() + "..."
    return cleaned or subject.strip()


def guess_subtitle(combined_text_lower):
    for hint in POSITION_HINTS:
        if hint in combined_text_lower:
            return hint.capitalize()
    return ""


def guess_category(combined_text_lower, domain):
    haystack = combined_text_lower + " " + domain.lower()
    for category, keywords in CATEGORY_RULES:
        if any(kw in haystack for kw in keywords):
            return category
    return "Sonstiges"


def guess_org_from_domain(domain):
    name = domain.split(".")[0]
    return name.replace("-", " ").replace("_", " ").title()


def fetch_existing(supabase_url, supabase_key, table, row_id):
    endpoint = supabase_url.rstrip("/") + f"/rest/v1/{table}?id=eq.{row_id}&select=data"
    req = urllib.request.Request(
        endpoint,
        method="GET",
        headers={
            "apikey": supabase_key,
            "Authorization": f"Bearer {supabase_key}",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            rows = json.loads(resp.read().decode("utf-8"))
            if rows:
                return rows[0]["data"]
    except Exception as e:
        print(f"Warnung: Konnte {table}/{row_id} nicht laden: {e}", file=sys.stderr)
    return None


def upsert_row(supabase_url, supabase_key, table, row_id, data):
    payload = {"id": row_id, "data": data}
    endpoint = supabase_url.rstrip("/") + f"/rest/v1/{table}?on_conflict=id"
    req = urllib.request.Request(
        endpoint,
        data=json.dumps(payload).encode("utf-8"),
        method="POST",
        headers={
            "apikey": supabase_key,
            "Authorization": f"Bearer {supabase_key}",
            "Content-Type": "application/json",
            "Prefer": "resolution=merge-duplicates,return=minimal",
        },
    )
    try:
        urllib.request.urlopen(req, timeout=15)
    except Exception as e:
        print(f"Fehler beim Speichern von {row_id} in {table}: {e}", file=sys.stderr)


def main():
    supabase_url = os.environ["SUPABASE_URL"]
    supabase_key = os.environ["SUPABASE_ANON_KEY"]

    conn = connect()
    try:
        folder = select_sent_folder(conn)
        print(f"Nutze Ordner: {folder}")

        messages = fetch_recent_messages(conn, days=2)
        print(f"{len(messages)} Nachricht(en) der letzten 2 Tage gefunden.")

        processed = 0
        for msg in messages:
            subject = decode_str(msg.get("Subject"))
            body = extract_body_text(msg)
            combined_text = subject + "\n" + body
            combined_lower = combined_text.lower()

            if not matches_keywords(combined_text):
                continue

            to_header = msg.get("To", "")
            display_name, addr = parseaddr(to_header)
            recipient = addr or to_header
            domain = recipient.split("@")[-1].lower() if "@" in recipient else "unbekannt"

            try:
                dt = parsedate_to_datetime(msg.get("Date"))
                iso_date = dt.date().isoformat()
            except Exception:
                iso_date = ""

            headline = clean_headline(subject)
            subtitle = guess_subtitle(combined_lower)
            bereich = guess_category(combined_lower, domain)
            org = guess_org_from_domain(domain)
            person = display_name or ""

            card_id = "org-" + hashlib.sha1(domain.encode("utf-8")).hexdigest()[:16]
            existing = fetch_existing(supabase_url, supabase_key, "bewerbungen", card_id)

            mail_entry = {"datum": iso_date, "betreff": subject}

            if existing:
                mailverlauf = existing.get("mailverlauf", [])
                already_logged = any(
                    e.get("datum") == mail_entry["datum"] and e.get("betreff") == mail_entry["betreff"]
                    for e in mailverlauf
                )
                if not already_logged:
                    mailverlauf.append(mail_entry)
                new_data = dict(existing)
                new_data["mailverlauf"] = mailverlauf
                new_data["datum"] = iso_date
                if person:
                    new_data["person"] = person
                new_data.setdefault("org", org)
                new_data.setdefault("headline", headline)
                new_data.setdefault("subtitle", subtitle)
                new_data.setdefault("bereich", bereich)
                new_data.setdefault("status", "yellow")
                new_data.setdefault("mail", recipient)
                new_data.setdefault("link", "")
                new_data.setdefault("notes", "")
            else:
                new_data = {
                    "org": org,
                    "headline": headline,
                    "subtitle": subtitle,
                    "bereich": bereich,
                    "person": person,
                    "mail": recipient,
                    "link": "",
                    "datum": iso_date,
                    "status": "yellow",
                    "notes": "",
                    "mailverlauf": [mail_entry],
                }

            upsert_row(supabase_url, supabase_key, "bewerbungen", card_id, new_data)

            # Netzwerk-Kontakt pflegen
            contact_id = "contact-" + hashlib.sha1(recipient.encode("utf-8")).hexdigest()[:16]
            existing_contact = fetch_existing(supabase_url, supabase_key, "kontakte", contact_id)
            if existing_contact:
                contact_data = dict(existing_contact)
                contact_data["letzter_kontakt"] = iso_date
                if person:
                    contact_data["name"] = person
                contact_data.setdefault("unternehmen", org)
                contact_data.setdefault("notizen", "")
            else:
                contact_data = {
                    "name": person or (recipient.split("@")[0] if "@" in recipient else recipient),
                    "email": recipient,
                    "unternehmen": org,
                    "erster_kontakt": iso_date,
                    "letzter_kontakt": iso_date,
                    "notizen": "",
                }
            upsert_row(supabase_url, supabase_key, "kontakte", contact_id, contact_data)

            processed += 1

        print(f"{processed} passende Bewerbungs-Mail(s) verarbeitet.")
    finally:
        conn.logout()


if __name__ == "__main__":
    main()
