"""
Bewerbungs-Mail-Sync (Cloud-Version, regelbasiert - keine bezahlte KI-API)

Ablauf pro Lauf:
  1. Sent-Ordner der letzten SENT_DAYS Tage nach Bewerbungs-Mails durchsuchen
     (Stichwoerter in Betreff oder Text) und pro FIRMA (Empfaenger-Domain)
     genau einen Datenpunkt in der Tabelle "bewerbungen" anlegen/pflegen.
  2. Fuer alle bekannten Firmen den Posteingang der letzten INBOX_DAYS Tage
     nach Antworten durchsuchen und diese dem Mailverlauf hinzufuegen.
  3. Status aus dem Inhalt der Antwort ableiten
     (offen / abgelehnt / angenommen / talent_pool).
  4. Kurze Zusammenfassung des Verlaufs in die Notizen schreiben.
  5. Pro Firma einen Kontakt in der Tabelle "kontakte" pflegen; dessen
     Bereich folgt immer dem Bereich der Bewerbung.

Benoetigte Umgebungsvariablen (als GitHub Secrets):
  ICLOUD_EMAIL, ICLOUD_APP_PASSWORD, SUPABASE_URL, SUPABASE_ANON_KEY

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

SENT_DAYS = 2      # wie weit zurueck neue Bewerbungen erkannt werden
INBOX_DAYS = 21    # wie weit zurueck nach Antworten gesucht wird

MAX_BODY_CHARS = 20000

KEYWORDS = ["bewerbung", "anschreiben", "application", "praktikum", "initiativbewerbung"]

CANDIDATE_SENT_FOLDERS = ["Sent Messages", "Sent", "INBOX.Sent Messages"]

CATEGORIES = [
    "Journalismus",
    "Wissenschaft",
    "Start-up",
    "Gründungszentrum",
    "KI",
    "Sonstiges",
]

CATEGORY_RULES = [
    ("Journalismus", ["redaktion", "journalismus", "journalist", "presse", "rundfunk",
                      "podcast", "nachrichten", "reporter", "volontariat", "medien"]),
    ("KI", ["künstliche intelligenz", "kuenstliche intelligenz", " ki ", "ki-",
            "artificial intelligence", "machine learning", " ai ", "llm", "deep learning"]),
    ("Gründungszentrum", ["stiftung", "accelerator", "inkubator", "incubator",
                          "gründungszentrum", "gruendungszentrum", "entrepreneurship",
                          "gründerzentrum", "venture", "start-hub"]),
    ("Wissenschaft", ["universität", "universitaet", "uni-", "hochschule", "forschung",
                      "wissenschaftlich", "lehrstuhl", "institut", "promotion", "fakultät"]),
    ("Start-up", ["start-up", "startup", "founder", "gründer", "gruender",
                  "cfo", "coo", "cto", "associate"]),
]

POSITION_PATTERNS = [
    r"Associate\s+(?:to|for|at)\s+the\s+[\w/]+(?:\s+[\w/]+)?",
    r"Praktikum(?:s)?\s*(?:als\s+)?[A-ZÄÖÜ][\wÄÖÜäöüß]+(?:\s+[A-ZÄÖÜ][\wÄÖÜäöüß]+)?",
    r"Werkstudent(?:in)?\s*(?:\([^)]*\))?",
    r"Kursleiter(?:in)?",
    r"Volontariat",
    r"Projekt[- ]?Manager(?:in)?",
    r"Programm[- ]?Manager(?:in)?",
    r"Wissenschaftliche(?:r)?\s+Mitarbeiter(?:in)?",
    r"Trainee(?:\s+[A-ZÄÖÜ][\wÄÖÜäöüß]+)?",
]

# ---------------------------------------------------------------- Status

STATUS_OFFEN = "offen"
STATUS_ABGELEHNT = "abgelehnt"
STATUS_ANGENOMMEN = "angenommen"
STATUS_TALENTPOOL = "talent_pool"

TALENTPOOL_KEYWORDS = [
    "talent pool", "talentpool", "talent-pool", "evidenz", "vormerken",
    "vorgemerkt", "in erinnerung behalten", "melden uns bei passenden",
    "keep your profile", "future opportunities",
]
ABLEHNUNG_KEYWORDS = [
    "leider", "absage", "nicht berücksichtigen", "nicht weiter",
    "andere kandidat", "unfortunately", "we regret", "not moving forward",
    "another candidate", "nicht zum zuge",
]
ZUSAGE_KEYWORDS = [
    "zusage", "freuen uns", "einladen", "einladung", "kennenlernen",
    "vorstellungsgespräch", "vorstellungsgespraech", "gespräch", "gespraech",
    "interview", "termin vereinbaren", "kennenzulernen",
    "we would like to invite", "happy to invite", "schedule a call",
    "next round", "naechste runde", "nächste runde",
]

# ------------------------------------------------- Ansprechperson-Erkennung

ANREDE_PATTERNS = [
    r"Sehr\s+geehrte(?:r)?\s+(?:Frau|Herr)\s+((?:Dr\.|Prof\.)?\s*[A-ZÄÖÜ][\wÄÖÜäöüß'-]+(?:\s+[A-ZÄÖÜ][\wÄÖÜäöüß'-]+)?)",
    r"(?:Hallo|Liebe|Lieber|Guten\s+Tag)\s+(?:Frau|Herr)\s+([A-ZÄÖÜ][\wÄÖÜäöüß'-]+(?:\s+[A-ZÄÖÜ][\wÄÖÜäöüß'-]+)?)",
    r"(?:Hallo|Liebe|Lieber|Hi|Hey|Dear)\s+([A-ZÄÖÜ][\wÄÖÜäöüß'-]{2,}(?:\s+[A-ZÄÖÜ][\wÄÖÜäöüß'-]+)?)\s*[,!\n]",
]

REPLY_HEADER_PATTERNS = [
    r"(?:^|\n)\s*>?\s*Von:\s*([^<\n\r]+?)\s*(?:<|\n|\r|$)",
    r"(?:^|\n)\s*>?\s*From:\s*([^<\n\r]+?)\s*(?:<|\n|\r|$)",
]

GREETING_PATTERNS = [
    r"(?:Viele|Beste|Herzliche|Freundliche|Liebe)\s+Gr(?:ü|ue)(?:ß|ss)e[,!]?",
    r"Mit\s+freundlichen\s+Gr(?:ü|ue)(?:ß|ss)en[,!]?",
    r"(?:Best|Kind|Warm)\s+regards[,!]?",
    r"(?:Cheers|Thanks|Danke|LG|VG|BG)[,!]",
]

NAME_STOPWORDS = [
    "gmbh", "ag", "e.v.", "team", "redaktion", "www.", "http", "@", "tel",
    "mobil", "fon", "fax", "str.", "strasse", "straße", "gesendet", "sent",
    "damen", "herren", "sir", "madam",
]

GENERIC_MAILBOXES = {
    "jobs", "job", "info", "kontakt", "contact", "praktikum", "praktika",
    "bewerbung", "bewerbungen", "career", "careers", "hr", "office", "mail",
    "team", "noreply", "no-reply", "hallo", "hello", "service", "support",
    "presse", "press", "redaktion", "post", "buero", "admin", "recruiting",
}


def looks_like_name(line):
    s = (line or "").strip(" >-*\t")
    if not (2 <= len(s) <= 40):
        return False
    low = s.lower()
    if any(w in low for w in NAME_STOPWORDS):
        return False
    words = s.split()
    if not (1 <= len(words) <= 4):
        return False
    return any(w[:1].isupper() for w in words if w)


def extract_person(body, to_display_name, recipient):
    """Ansprechperson: Anrede -> An-Feld -> Antwort-Header -> Signatur -> Adresse."""
    # 1. Anrede in der eigenen Mail (nicht im zitierten Teil)
    own_part = body.split("\n>")[0]
    for pattern in ANREDE_PATTERNS:
        m = re.search(pattern, own_part)
        if m:
            cand = m.group(1).strip()
            if looks_like_name(cand):
                return cand

    # 2. Anzeigename im An-Feld
    if to_display_name and looks_like_name(to_display_name):
        return to_display_name.strip()

    # 3. Antwort-Header im zitierten Teil
    for pattern in REPLY_HEADER_PATTERNS:
        m = re.search(pattern, body, re.IGNORECASE)
        if m:
            cand = m.group(1).strip().strip('"')
            if looks_like_name(cand):
                return cand

    m = re.search(r"([^<\n\r,]{2,60}?)\s*<[^>]+>\s*(?:wrote|schrieb)\s*:", body, re.IGNORECASE)
    if m:
        words = [w for w in m.group(1).strip().strip('">').split() if w[:1].isupper()]
        cand = " ".join(words[-3:]) if words else ""
        if looks_like_name(cand):
            return cand

    # 4. Signatur im zitierten Abschnitt
    lines = body.splitlines()
    for i, line in enumerate(lines):
        for gp in GREETING_PATTERNS:
            if re.search(r"^\s*>+\s*" + gp + r"\s*$", line, re.IGNORECASE):
                for nxt in lines[i + 1:i + 4]:
                    if looks_like_name(nxt):
                        return nxt.strip(" >-*\t")
                break

    # 5. Aus der Mailadresse ableiten
    local = recipient.split("@")[0] if "@" in recipient else ""
    if local.lower() in GENERIC_MAILBOXES:
        return ""
    if local and not local.isdigit():
        parts = [p for p in re.split(r"[._-]+", local) if p.isalpha() and len(p) > 1]
        if 1 <= len(parts) <= 3:
            cand = " ".join(p.capitalize() for p in parts)
            if looks_like_name(cand):
                return cand
    return ""


# ----------------------------------------------------------------- Helpers

def decode_str(raw):
    if raw is None:
        return ""
    out = ""
    for text, enc in decode_header(raw):
        if isinstance(text, bytes):
            out += text.decode(enc or "utf-8", errors="replace")
        else:
            out += text
    return out


def extract_body_text(msg):
    parts = []
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain" and \
               "attachment" not in str(part.get("Content-Disposition", "")).lower():
                try:
                    payload = part.get_payload(decode=True)
                    if payload:
                        parts.append(payload.decode(part.get_content_charset() or "utf-8",
                                                    errors="replace"))
                except Exception:
                    continue
    else:
        try:
            payload = msg.get_payload(decode=True)
            if payload:
                parts.append(payload.decode(msg.get_content_charset() or "utf-8",
                                            errors="replace"))
        except Exception:
            pass
    return "\n".join(parts)[:MAX_BODY_CHARS]


def matches_keywords(text):
    low = text.lower()
    return any(kw in low for kw in KEYWORDS)


def guess_subtitle(text):
    for pattern in POSITION_PATTERNS:
        m = re.search(pattern, text)
        if m:
            cand = m.group(0).strip()
            return cand[:42].rstrip() + "..." if len(cand) > 45 else cand
    return ""


def guess_category(text_lower, domain):
    hay = text_lower + " " + domain.lower()
    for category, keywords in CATEGORY_RULES:
        if any(kw in hay for kw in keywords):
            return category
    return "Sonstiges"


def guess_org_from_domain(domain):
    return domain.split(".")[0].replace("-", " ").replace("_", " ").title()


def classify_reply(text):
    """Leitet den Status aus dem Text einer Antwort ab."""
    low = text.lower()
    if any(k in low for k in TALENTPOOL_KEYWORDS):
        return STATUS_TALENTPOOL
    if any(k in low for k in ABLEHNUNG_KEYWORDS):
        return STATUS_ABGELEHNT
    if any(k in low for k in ZUSAGE_KEYWORDS):
        return STATUS_ANGENOMMEN
    return None


def fmt_de(iso_date):
    try:
        y, m, d = iso_date.split("-")
        return f"{d}.{m}.{y}"
    except Exception:
        return iso_date or ""


def build_summary(mailverlauf, status):
    """Kurze Zusammenfassung des Verlaufs fuer das Notizfeld."""
    if not mailverlauf:
        return ""
    entries = sorted(mailverlauf, key=lambda e: e.get("datum", ""))
    raus = sum(1 for e in entries if e.get("richtung") == "ausgehend")
    rein = sum(1 for e in entries if e.get("richtung") == "eingehend")

    teile = []
    first = entries[0]
    teile.append(f"Erstkontakt {fmt_de(first.get('datum',''))}: {first.get('betreff','')}".strip())
    if rein:
        letzte_antwort = [e for e in entries if e.get("richtung") == "eingehend"][-1]
        teile.append(f"Letzte Antwort {fmt_de(letzte_antwort.get('datum',''))}: "
                     f"{letzte_antwort.get('betreff','')}".strip())
    else:
        teile.append("Bisher keine Antwort erhalten")

    status_text = {
        STATUS_OFFEN: "Status offen",
        STATUS_ABGELEHNT: "Absage erhalten",
        STATUS_ANGENOMMEN: "Positive Rueckmeldung",
        STATUS_TALENTPOOL: "In den Talent Pool aufgenommen",
    }.get(status, "Status offen")

    teile.append(f"{raus} gesendet / {rein} erhalten - {status_text}")
    return " | ".join(t for t in teile if t)


# ---------------------------------------------------------------- IMAP

def connect():
    conn = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    conn.login(os.environ["ICLOUD_EMAIL"], os.environ["ICLOUD_APP_PASSWORD"])
    return conn


def list_folder_names(conn):
    typ, mailboxes = conn.list()
    names = []
    if typ == "OK":
        for m in mailboxes:
            match = re.search(r'"([^"]+)"$', m.decode(errors="replace"))
            if match:
                names.append(match.group(1))
    return names


def select_sent_folder(conn):
    names = list_folder_names(conn)
    for cand in CANDIDATE_SENT_FOLDERS:
        if cand in names and conn.select(f'"{cand}"', readonly=True)[0] == "OK":
            return cand
    for name in names:
        if "sent" in name.lower() and conn.select(f'"{name}"', readonly=True)[0] == "OK":
            return name
    raise RuntimeError("Kein Sent-Ordner gefunden. Gefunden: " + ", ".join(names))


def fetch_since(conn, days):
    since = (datetime.date.today() - datetime.timedelta(days=days)).strftime("%d-%b-%Y")
    typ, data = conn.search(None, f"(SINCE {since})")
    if typ != "OK" or not data or data[0] is None:
        return []
    out = []
    for mid in data[0].split():
        typ, md = conn.fetch(mid, "(RFC822)")
        if typ == "OK" and md and md[0]:
            out.append(email.message_from_bytes(md[0][1]))
    return out


def msg_date(msg):
    try:
        return parsedate_to_datetime(msg.get("Date")).date().isoformat()
    except Exception:
        return ""


# ------------------------------------------------------------- Supabase

def sb_get(url, key, table, row_id=None):
    endpoint = url.rstrip("/") + f"/rest/v1/{table}?select=id,data"
    if row_id:
        endpoint += f"&id=eq.{row_id}"
    req = urllib.request.Request(endpoint, method="GET",
                                 headers={"apikey": key, "Authorization": f"Bearer {key}"})
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        print(f"Warnung: {table} konnte nicht geladen werden: {e}", file=sys.stderr)
        return []


def sb_upsert(url, key, table, row_id, data):
    payload = {"id": row_id, "data": data}
    endpoint = url.rstrip("/") + f"/rest/v1/{table}?on_conflict=id"
    req = urllib.request.Request(
        endpoint, data=json.dumps(payload).encode("utf-8"), method="POST",
        headers={"apikey": key, "Authorization": f"Bearer {key}",
                 "Content-Type": "application/json",
                 "Prefer": "resolution=merge-duplicates,return=minimal"})
    try:
        urllib.request.urlopen(req, timeout=20)
    except Exception as e:
        print(f"Fehler beim Speichern von {row_id} in {table}: {e}", file=sys.stderr)


def card_id_for(domain):
    return "org-" + hashlib.sha1(domain.encode("utf-8")).hexdigest()[:16]


def contact_id_for(domain):
    return "contact-" + hashlib.sha1(domain.encode("utf-8")).hexdigest()[:16]


def add_mail_entry(mailverlauf, entry):
    """Fuegt einen Eintrag hinzu, falls er noch nicht existiert."""
    for e in mailverlauf:
        if e.get("datum") == entry["datum"] and e.get("betreff") == entry["betreff"] \
           and e.get("richtung") == entry["richtung"]:
            return False
    mailverlauf.append(entry)
    return True


# ------------------------------------------------------------------ Main

def main():
    url = os.environ["SUPABASE_URL"]
    key = os.environ["SUPABASE_ANON_KEY"]

    conn = connect()
    cards = {}   # domain -> data dict

    try:
        # ---------- 1. Gesendete Bewerbungen ----------
        folder = select_sent_folder(conn)
        print(f"Sent-Ordner: {folder}")
        sent = fetch_since(conn, SENT_DAYS)
        print(f"{len(sent)} gesendete Nachricht(en) der letzten {SENT_DAYS} Tage.")

        for msg in sent:
            subject = decode_str(msg.get("Subject"))
            body = extract_body_text(msg)
            combined = subject + "\n" + body
            if not matches_keywords(combined):
                continue

            display_name, addr = parseaddr(msg.get("To", ""))
            recipient = addr or msg.get("To", "")
            if "@" not in recipient:
                continue
            domain = recipient.split("@")[-1].lower()
            iso = msg_date(msg)

            existing_rows = sb_get(url, key, "bewerbungen", card_id_for(domain))
            data = dict(existing_rows[0]["data"]) if existing_rows else {}

            data.setdefault("org", guess_org_from_domain(domain))
            data.setdefault("bereich", guess_category(combined.lower(), domain))
            data.setdefault("status", STATUS_OFFEN)
            data.setdefault("link", "")
            data.setdefault("mailverlauf", [])
            if not data.get("subtitle"):
                data["subtitle"] = guess_subtitle(combined)
            person = extract_person(body, display_name, recipient)
            if person:
                data["person"] = person
            data.setdefault("person", "")
            data["mail"] = recipient
            data["domain"] = domain

            add_mail_entry(data["mailverlauf"],
                           {"datum": iso, "betreff": subject, "richtung": "ausgehend"})
            cards[domain] = data

        # ---------- 2. Bekannte Firmen aus der Datenbank ergaenzen ----------
        for row in sb_get(url, key, "bewerbungen"):
            d = row["data"]
            dom = d.get("domain") or (d.get("mail", "").split("@")[-1].lower()
                                      if "@" in d.get("mail", "") else None)
            if dom and dom not in cards:
                d = dict(d)
                d.setdefault("mailverlauf", [])
                d["domain"] = dom
                cards[dom] = d

        # ---------- 3. Antworten im Posteingang suchen ----------
        if cards:
            conn.select("INBOX", readonly=True)
            inbox = fetch_since(conn, INBOX_DAYS)
            print(f"{len(inbox)} Nachricht(en) im Posteingang der letzten {INBOX_DAYS} Tage.")

            for msg in inbox:
                _, from_addr = parseaddr(msg.get("From", ""))
                if "@" not in from_addr:
                    continue
                dom = from_addr.split("@")[-1].lower()
                if dom not in cards:
                    continue

                subject = decode_str(msg.get("Subject"))
                body = extract_body_text(msg)
                iso = msg_date(msg)
                data = cards[dom]

                neu = add_mail_entry(data["mailverlauf"],
                                     {"datum": iso, "betreff": subject, "richtung": "eingehend"})
                if neu:
                    print(f"  Antwort von {dom}: {subject[:60]}")

                erkannt = classify_reply(subject + "\n" + body)
                if erkannt:
                    data["status"] = erkannt

        # ---------- 4. Speichern ----------
        for domain, data in cards.items():
            data["mailverlauf"] = sorted(data["mailverlauf"], key=lambda e: e.get("datum", ""))
            if data["mailverlauf"]:
                data["datum"] = data["mailverlauf"][-1].get("datum", "")
            data["notes"] = build_summary(data["mailverlauf"], data.get("status", STATUS_OFFEN))

            sb_upsert(url, key, "bewerbungen", card_id_for(domain), data)

            if data.get("person") or data.get("mail"):
                contact_rows = sb_get(url, key, "kontakte", contact_id_for(domain))
                cd = dict(contact_rows[0]["data"]) if contact_rows else {}
                cd["name"] = data.get("person") or cd.get("name") or ""
                cd["email"] = data.get("mail", "")
                cd["unternehmen"] = data.get("org", "")
                cd["bereich"] = data.get("bereich", "Sonstiges")
                cd.setdefault("notizen", "")
                if data["mailverlauf"]:
                    cd.setdefault("erster_kontakt", data["mailverlauf"][0].get("datum", ""))
                    cd["letzter_kontakt"] = data["mailverlauf"][-1].get("datum", "")
                sb_upsert(url, key, "kontakte", contact_id_for(domain), cd)

        print(f"{len(cards)} Firma/Firmen verarbeitet.")
    finally:
        conn.logout()


if __name__ == "__main__":
    main()
