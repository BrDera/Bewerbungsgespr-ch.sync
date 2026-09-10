"""
Bewerbungs-Mail-Sync (Cloud-Version, regelbasiert - keine bezahlte KI-API)

Durchsucht per IMAP den Sent-Ordner einer iCloud-Adresse nach Mails der
letzten 2 Tage, filtert nach Stichwörtern im Betreff und im Text,
gruppiert Treffer PRO FIRMA (per Empfänger-Domain) und pflegt dafür
jeweils genau eine Zeile in der Supabase-Tabelle "bewerbungen" inkl.
Mailverlauf. Außerdem wird pro Empfänger ein Eintrag in der Tabelle
"kontakte" gepflegt (für den Netzwerk-Tab im Board).

Gedacht zum Ausführen in GitHub Actions nach Zeitplan - siehe
mail-sync.yml. Benötigte Umgebungsvariablen (als GitHub Secrets):
  ICLOUD_EMAIL         volle iCloud-Adresse, z.B. name@icloud.com
  ICLOUD_APP_PASSWORD  App-spezifisches Passwort von appleid.apple.com
  SUPABASE_URL         z.B. https://xxxx.supabase.co
  SUPABASE_ANON_KEY    der "anon public" Key aus Supabase

Nutzt nur die Python-Standardbibliothek, kein pip install nötig.
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

# Stichwörter, die im Betreff ODER im Text auf eine Bewerbung hindeuten
# (klein geschrieben). Bei Bedarf ergänzen/anpassen.
KEYWORDS = ["bewerbung", "anschreiben", "application", "praktikum", "initiativbewerbung"]

CANDIDATE_SENT_FOLDERS = ["Sent Messages", "Sent", "INBOX.Sent Messages"]

MAX_BODY_CHARS = 20000

CATEGORIES = [
    "Journalismus",
    "Gründung/Gründungszentrum",
    "Wissenschaft/Akademisch",
    "Start-up",
    "KI",
    "Kurzfristig/Nebenjob",
    "Sonstiges",
]

# Regelbasierte Kategorie-Zuordnung: Schlüsselwort -> Kategorie.
# Wird gegen Betreff + Text + Empfänger-Domain geprüft, erster Treffer gewinnt.
CATEGORY_RULES = [
    ("Journalismus", ["redaktion", "journalismus", "journalist", "presse", "rundfunk", "podcast", "nachrichten"]),
    ("Start-up", ["start-up", "startup", "founder", "gründer", "gruender", "cfo", "coo", "cto"]),
    ("Gründung/Gründungszentrum", ["stiftung", "accelerator", "inkubator", "gründungszentrum", "gruendungszentrum", "entrepreneurship"]),
    ("Wissenschaft/Akademisch", ["universität", "universitaet", "uni ", "hochschule", "forschung", "wissenschaftlich", "lehrstuhl", "institut"]),
    ("KI", ["künstliche intelligenz", "kuenstliche intelligenz", " ki ", "artificial intelligence", "machine learning", " ai "]),
    ("Kurzfristig/Nebenjob", ["nebenjob", "aushilfe", "kellner", "kellnerin", "lieferant", "fahrer", "café", "cafe"]),
]

# Regex-Muster, die eine ganze Positions-/Rollenphrase einfangen (nicht nur
# ein einzelnes Wort), z.B. "Associate to the CFO/COO" oder "Praktikum Redaktion".
# Werden der Reihe nach gegen Betreff+Text (Original-Groß-/Kleinschreibung) geprüft.
POSITION_PATTERNS = [
    r"Associate\s+(?:to|for|at)\s+the\s+[\w/]+(?:\s+[\w/]+)?",
    r"Praktikum(?:s)?\s*(?:als\s+)?[A-ZÄÖÜ][\wÄÖÜäöüß]+(?:\s+[A-ZÄÖÜ][\wÄÖÜäöüß]+)?",
    r"Werkstudent(?:in)?\s*(?:\([^)]*\))?",
    r"Kursleiter(?:in)?",
    r"Volontariat",
    r"Projekt[- ]?Manager(?:in)?",
    r"Programm[- ]?Manager(?:in)?",
    r"Trainee(?:\s+[A-ZÄÖÜ][\wÄÖÜäöüß]+)?",
]


# Muster, um den Namen der Gegenseite aus zitierten Antwort-Teilen zu lesen.
# In gesendeten Mails steht die eigene Signatur oben - der Name des Gegenuebers
# taucht im zitierten Abschnitt auf (Antwort-Header oder dessen Signatur).
REPLY_HEADER_PATTERNS = [
    r"(?:^|\n)\s*>?\s*Von:\s*([^<\n\r]+?)\s*(?:<|\n|\r|$)",
    r"(?:^|\n)\s*>?\s*From:\s*([^<\n\r]+?)\s*(?:<|\n|\r|$)",
    r"Am\s+.{5,40}\s+schrieb\s+([^<\n\r:]+?)\s*(?:<|:|\n|$)",
]

# Grussformeln, nach denen in einer Signatur der Name folgt.
GREETING_PATTERNS = [
    r"(?:Viele|Beste|Herzliche|Freundliche|Liebe)\s+Gr(?:ü|ue)(?:ß|ss)e[,!]?",
    r"Mit\s+freundlichen\s+Gr(?:ü|ue)(?:ß|ss)en[,!]?",
    r"(?:Best|Kind|Warm)\s+regards[,!]?",
    r"(?:Cheers|Thanks|Danke|LG|VG|BG)[,!]",
]

# Zeilen, die nie ein Personenname sind (Signatur-Rauschen).
# Generische Postfaecher, aus denen kein Personenname abgeleitet werden darf.
GENERIC_MAILBOXES = {
    "jobs", "job", "info", "kontakt", "contact", "praktikum", "praktika",
    "bewerbung", "bewerbungen", "career", "careers", "hr", "office", "mail",
    "team", "noreply", "no-reply", "hallo", "hello", "service", "support",
    "presse", "press", "redaktion", "post", "buero", "admin", "recruiting",
}

NAME_STOPWORDS = [
    "gmbh", "ag", "e.v.", "team", "redaktion", "www.", "http", "@", "tel",
    "mobil", "fon", "fax", "str.", "strasse", "straße", "gesendet", "sent",
]


def looks_like_name(line):
    """Grobe Pruefung, ob eine Zeile ein Personenname sein koennte."""
    s = line.strip(" >-*\t")
    if not (2 <= len(s) <= 40):
        return False
    low = s.lower()
    if any(w in low for w in NAME_STOPWORDS):
        return False
    words = s.split()
    if not (1 <= len(words) <= 4):
        return False
    # mindestens ein Wort muss mit Grossbuchstaben beginnen
    return any(w[:1].isupper() for w in words if w)


def extract_person(body, to_display_name, recipient):
    """Ermittelt den Namen der Ansprechperson.

    Reihenfolge: Anzeigename im An-Feld -> Antwort-Header im zitierten Text
    -> Signatur (Name nach einer Grussformel) im zitierten Text
    -> Ableitung aus der Mailadresse.
    """
    if to_display_name and looks_like_name(to_display_name):
        return to_display_name.strip()

    for pattern in REPLY_HEADER_PATTERNS:
        m = re.search(pattern, body, re.IGNORECASE)
        if m:
            cand = m.group(1).strip().strip('"')
            if looks_like_name(cand):
                return cand

    # "... Amy Fischer <amy@x.de> wrote:" / "... schrieb Max Mustermann <m@x.de>:"
    m = re.search(r"([^<\n\r,]{2,60}?)\s*<[^>]+>\s*(?:wrote|schrieb)\s*:", body, re.IGNORECASE)
    if m:
        tail = m.group(1).strip().strip('">')
        words = [w for w in tail.split() if w[:1].isupper()]
        cand = " ".join(words[-3:]) if words else ""
        if looks_like_name(cand):
            return cand

    # Signatur im zitierten Abschnitt: Name direkt nach einer Grussformel
    lines = body.splitlines()
    for i, line in enumerate(lines):
        for gp in GREETING_PATTERNS:
            if re.search(r"^\s*>+\s*" + gp + r"\s*$", line, re.IGNORECASE):
                for nxt in lines[i + 1:i + 4]:
                    if looks_like_name(nxt):
                        return nxt.strip(" >-*\t")
                break

    # Fallback: aus der Mailadresse ableiten (vorname.nachname@...)
    local = recipient.split("@")[0] if "@" in recipient else ""
    if local.lower() in GENERIC_MAILBOXES:
        return ""
    if local and not local.isdigit():
        parts = re.split(r"[._-]+", local)
        parts = [p for p in parts if p.isalpha() and len(p) > 1]
        if 1 <= len(parts) <= 3:
            cand = " ".join(p.capitalize() for p in parts)
            if looks_like_name(cand):
                return cand

    return ""


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


def guess_subtitle(combined_text):
    """Sucht eine ganze Positions-/Rollenphrase im Original-Text (Groß-/
    Kleinschreibung erhalten), damit z.B. 'Associate to the CFO/COO' oder
    'Praktikum Redaktion' komplett erkannt werden, nicht nur ein Wort."""
    for pattern in POSITION_PATTERNS:
        match = re.search(pattern, combined_text)
        if match:
            candidate = match.group(0).strip()
            if len(candidate) > 45:
                candidate = candidate[:42].rstrip() + "..."
            return candidate
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

            subtitle = guess_subtitle(combined_text)
            bereich = guess_category(combined_lower, domain)
            org = guess_org_from_domain(domain)
            person = extract_person(body, display_name, recipient)

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
                new_data.setdefault("subtitle", subtitle)
                new_data.setdefault("bereich", bereich)
                new_data.setdefault("status", "yellow")
                new_data.setdefault("mail", recipient)
                new_data.setdefault("link", "")
                new_data.setdefault("notes", "")
            else:
                new_data = {
                    "org": org,
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
            # Der Bereich des Kontakts folgt immer dem der Bewerbung derselben Firma
            final_bereich = new_data["bereich"]

            if existing_contact:
                contact_data = dict(existing_contact)
                contact_data["letzter_kontakt"] = iso_date
                if person:
                    contact_data["name"] = person
                contact_data["bereich"] = final_bereich
                contact_data.setdefault("unternehmen", org)
                contact_data.setdefault("notizen", "")
            else:
                contact_data = {
                    "name": person or (recipient.split("@")[0] if "@" in recipient else recipient),
                    "email": recipient,
                    "unternehmen": org,
                    "bereich": final_bereich,
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
