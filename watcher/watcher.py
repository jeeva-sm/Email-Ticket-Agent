#!/usr/bin/env python3
# watcher/watcher.py
"""
Outlook watcher:
 - Polls the configured mailbox in desktop Outlook via COM
 - Filters unread messages conservatively (whitelist keywords / triggers)
 - Runs heuristic parser first; if low-confidence or trigger present, escalate to LLM
 - Creates tickets in DB and moves processed messages to a "ProcessedByAgent" subfolder
 - Sends an acknowledgement email (Reply() where possible) after ticket creation
"""

from pathlib import Path
import argparse
import sys
import os
import time
import json
import hashlib
import logging
from datetime import datetime, timedelta
from typing import Optional

# Ensure project root (one level up) is on sys.path so imports like `import load_env` work
PROJECT_ROOT = str(Path(__file__).resolve().parents[1])
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

# Load environment variables (expects a `load_env.py` that loads .env into os.environ)
try:
    import load_env  # project-specific module to load .env into environment
except Exception:
    # If not present, we silently continue; callers may have set env some other way
    pass

# -----------------------
# Logging
# -----------------------
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger("watcher")

# -----------------------
# win32com / pywin32 (Outlook)
# -----------------------
# Use explicit imports to avoid "win32com.client.client" confusion
try:
    import win32com.client as win32com_client
    from win32com.client import Dispatch
except Exception:
    win32com_client = None
    Dispatch = None

def _get_outlook_namespace():
    """
    Return the Outlook MAPI namespace. Raises RuntimeError if pywin32 is not available.
    """
    if Dispatch is None:
        raise RuntimeError(
            "pywin32 (win32com) not installed. Install it in your venv: pip install pywin32"
        )
    try:
        outlook_app = Dispatch("Outlook.Application")
        return outlook_app.GetNamespace("MAPI")
    except Exception as e:
        raise RuntimeError(f"Failed to get Outlook namespace: {e}") from e

# -----------------------
# Config (via environment / defaults)
# -----------------------
POLL_INTERVAL = int(os.getenv("WATCHER_POLL_INTERVAL", "900"))  # seconds; default 15 minutes
MAX_MAILS_PER_POLL = int(os.getenv("WATCHER_MAX_MAILS_PER_POLL", "10"))
OUTLOOK_ACCOUNT = os.getenv("OUTLOOK_ACCOUNT", None)
PROCESSED_FOLDER_NAME = os.getenv("WATCHER_PROCESSED_FOLDER", "ProcessedByAgent")
INBOX_FOLDER_NAME = os.getenv("MAILBOX", "Inbox")

HEURISTIC_CONF_THRESHOLD = float(os.getenv("HEURISTIC_CONF_THRESHOLD", "0.75"))
LLM_CONF_THRESHOLD = float(os.getenv("LLM_CONF_THRESHOLD", "0.8"))

PROCESSED_FOLDER_NAME = os.getenv("WATCHER_PROCESSED_FOLDER", PROCESSED_FOLDER_NAME)
SYSADMIN_TRIGGERS = [
    s.strip().lower()
    for s in os.getenv(
        "SYSADMIN_TRIGGERS",
        "sysadmin,to sysadmin,sys admin team,sysadmin team,hi sysadmin,hi sys admin"
    ).split(",")
    if s.strip()
]

# Ping whitelist keywords (body/subject) that indicate IT/sysadmin work
WHITELIST_KEYWORDS = [k.strip().lower() for k in os.getenv("WATCHER_WHITELIST_KEYWORDS",
    "vpn,anydesk,printer,network,internet,password,access,openvpn,install,error,issue,failed,connect,remote,anydesk id"
).split(",") if k.strip()]

# Sender patterns to ignore
IGNORE_SENDERS = [s.strip().lower() for s in os.getenv("WATCHER_IGNORE_SENDERS",
    "noreply@,no-reply@,notifications@,do-not-reply@,bounce@").split(",") if s.strip()]

SUBJECT_BLACKLIST_REGEX = os.getenv("WATCHER_SUBJECT_BLACKLIST_REGEX",
    r"\b(birthday|invitation|calendar|newsletter|unsubscribe|social|celebration|party|invited|RSVP|out of office)\b"
)
MAX_AGE_DAYS = int(os.getenv("WATCHER_MAX_AGE_DAYS", "30"))
MAX_RECIPIENTS = int(os.getenv("WATCHER_MAX_RECIPIENTS", "50"))

# Ack / reply env
REPLY_ENABLED = os.getenv("REPLY_ENABLED", "true").lower() in ("1", "true", "yes")
REPLY_CC = os.getenv("REPLY_CC", "").strip()
REPLY_SUBJECT_TEMPLATE = os.getenv("REPLY_SUBJECT_TEMPLATE", "Re: {subject} - Ticket #{ticket_id}")
REPLY_BODY_TEMPLATE = os.getenv(
    "REPLY_BODY_TEMPLATE",
    "Hello {reporter_name_or_email},\n\n"
    "Thank you — we have created ticket #{ticket_id} for your request:\n"
    "Subject: {subject}\n"
    "Priority: {priority}\n\n"
    "Our team will review and update the ticket. If urgent, please reply with URGENT.\n\n"
    "Thanks,\nSysAdmin Team"
)

# LLM cache dir
LLM_CACHE_DIR = os.path.join(PROJECT_ROOT, ".cache_llm")
os.makedirs(LLM_CACHE_DIR, exist_ok=True)

# -----------------------
# Imports that depend on project layout
# -----------------------
# Import locally inside functions where possible to avoid import-time errors when running tests
# But import top-level parser/db modules used frequently
try:
    from parser.heuristic import parse_email_message as heuristic_parser
except Exception:
    heuristic_parser = None

try:
    from parser.llm_provider import parse_email_with_llm
except Exception:
    parse_email_with_llm = None

# DB models
try:
    from db.models import get_session, Ticket
except Exception:
    get_session = None
    Ticket = None

# -----------------------
# Helpers: LLM cache
# -----------------------
def _llm_cache_get(key: str):
    p = os.path.join(LLM_CACHE_DIR, hashlib.sha256(key.encode()).hexdigest() + ".json")
    if os.path.exists(p):
        try:
            return json.load(open(p, "r", encoding="utf-8"))
        except Exception:
            return None
    return None

def _llm_cache_set(key: str, value: dict):
    p = os.path.join(LLM_CACHE_DIR, hashlib.sha256(key.encode()).hexdigest() + ".json")
    json.dump(value, open(p, "w", encoding="utf-8"), indent=2)

# -----------------------
# Message / trigger helpers
# -----------------------
import re
import difflib
from html import unescape
import html as _html

_SUBJECT_BLACKLIST_RE = re.compile(SUBJECT_BLACKLIST_REGEX, flags=re.IGNORECASE)
FUZZY_THRESHOLD = float(os.getenv("SYSADMIN_FUZZY_THRESHOLD", "0.78"))

def _normalize_text_for_trigger(text: str) -> str:
    if not text:
        return ""
    s = unescape(text)
    s = re.sub(r"<[^>]+>", " ", s)  # strip HTML tags roughly
    s = re.sub(r"\s+", " ", s)
    return s.strip().lower()

def _tokens_from_text(text: str):
    return re.findall(r"\w+", text or "")

def _contains_trigger(texts: list) -> bool:
    combined = " ".join(_normalize_text_for_trigger(t or "") for t in texts)
    if not combined:
        return False

    # exact match for configured triggers
    for trig in SYSADMIN_TRIGGERS:
        trig = trig.strip().lower()
        if not trig:
            continue
        patt = r"\b" + re.escape(trig) + r"\b"
        if re.search(patt, combined):
            logger.debug("Exact trigger match: %r", trig)
            return True

    # fuzzy checks token-by-token and token pairs
    tokens = _tokens_from_text(combined)
    if not tokens:
        return False

    trig_forms = []
    for trig in SYSADMIN_TRIGGERS:
        t = trig.strip().lower()
        if not t:
            continue
        trig_forms.append(t)
        trig_forms.append(t.replace(" ", ""))

    for tok in tokens:
        for tf in trig_forms:
            if tok == tf:
                logger.debug("Token exact match fallback: %r == %r", tok, tf)
                return True
            ratio = difflib.SequenceMatcher(None, tok, tf).ratio()
            if ratio >= FUZZY_THRESHOLD:
                logger.info("Fuzzy trigger match: token=%r trig=%r ratio=%.2f", tok, tf, ratio)
                return True

    # check adjacent token pairs
    if len(tokens) >= 2:
        for i in range(len(tokens) - 1):
            pair = tokens[i] + tokens[i+1]
            for tf in trig_forms:
                ratio = difflib.SequenceMatcher(None, pair, tf).ratio()
                if ratio >= FUZZY_THRESHOLD:
                    logger.info("Fuzzy trigger match (pair): pair=%r trig=%r ratio=%.2f", pair, tf, ratio)
                    return True

    return False

# subject/body checks for conservative processing
def _extract_sender_email(msg):
    try:
        s = getattr(msg, "SenderEmailAddress", None)
        if s:
            return s.lower()
    except Exception:
        pass
    try:
        s = getattr(msg, "SentOnBehalfOfName", None)
        if s:
            return s.lower()
    except Exception:
        pass
    try:
        sender = getattr(msg, "Sender", None)
        if sender and getattr(sender, "Address", None):
            return sender.Address.lower()
    except Exception:
        pass
    return None

def _message_text(msg):
    try:
        subject = getattr(msg, "Subject", "") or ""
    except Exception:
        subject = ""
    try:
        body = getattr(msg, "Body", "") or ""
    except Exception:
        body = ""
    return subject, body

def should_process_message(msg) -> (bool, str):
    """
    Conservative filter:
      - Skip non-mail items
      - Skip empty subject/body
      - Skip blacklisted senders
      - Skip subject blacklisted (birthday/newsletter)
      - Skip too old
      - Skip mailing-lists (many recipients)
      - Only process if whitelist keyword present
    """
    try:
        mc = getattr(msg, "MessageClass", "") or ""
        if mc and not mc.lower().startswith("ipm.note"):
            return False, f"non-mail item (MessageClass={mc})"
    except Exception:
        pass

    subject = (getattr(msg, "Subject", "") or "").strip()
    body = (getattr(msg, "Body", "") or "").strip()
    if not subject and not body:
        return False, "empty subject/body"

    sender_email = _extract_sender_email(msg)
    if sender_email:
        for pat in IGNORE_SENDERS:
            if pat and pat in sender_email:
                return False, f"sender blacklisted ({pat})"

    if subject and _SUBJECT_BLACKLIST_RE.search(subject):
        return False, "subject blacklisted"

    try:
        recvd = getattr(msg, "ReceivedTime", None)
        if recvd and MAX_AGE_DAYS > 0:
            if isinstance(recvd, datetime):
                if datetime.utcnow() - recvd > timedelta(days=MAX_AGE_DAYS):
                    return False, f"too old (> {MAX_AGE_DAYS} days)"
    except Exception:
        pass

    try:
        recips = getattr(msg, "Recipients", None)
        if recips:
            count = getattr(recips, "Count", None)
            if count and MAX_RECIPIENTS > 0 and int(count) > MAX_RECIPIENTS:
                return False, f"mailing list too large (recipients={count})"
    except Exception:
        pass

    s_lower = subject.lower()
    body_lower = body.lower()
    for kw in WHITELIST_KEYWORDS:
        if kw and (kw in s_lower or kw in body_lower):
            return True, f"contains keyword '{kw}'"

    # fallback: if trigger term present anywhere (subject/body/to/cc/html), allow processing
    html_body = ""
    try:
        html_body = getattr(msg, "HTMLBody", "") or ""
    except Exception:
        html_body = ""
    to_field = getattr(msg, "To", "") or ""
    cc_field = getattr(msg, "CC", "") or ""
    if _contains_trigger([subject, body, html_body, to_field, cc_field]):
        return True, "trigger word present"

    return False, "no ticket keywords found"

# -----------------------
# Acknowledgement / auto-reply helpers
# -----------------------
def _simple_email_valid(addr: str) -> bool:
    if not addr:
        return False
    return bool(re.match(r"^[^@ \t\r\n]+@[^@ \t\r\n]+\.[^@ \t\r\n]+$", addr))

def _is_placeholder_domain(addr: str) -> bool:
    if not addr:
        return True
    low = addr.lower()
    return any(low.endswith(d) for d in ("@example.com", "@localhost", "@invalid", "@example.net"))

def _filter_cc_list(cc_raw: str) -> list:
    if not cc_raw:
        return []
    parts = re.split(r"[;,]", cc_raw)
    out = []
    for p in parts:
        a = p.strip()
        if not a:
            continue
        if _is_placeholder_domain(a):
            logger.debug("Skipping placeholder CC: %s", a)
            continue
        if _simple_email_valid(a):
            out.append(a)
        else:
            logger.debug("Skipping invalid-format CC: %s", a)
    return out

def _safe_get_reporter_name(mail):
    try:
        name = getattr(mail, "SenderName", None)
        if name:
            return name
    except Exception:
        pass
    try:
        return getattr(mail, "SenderEmailAddress", None) or "requester"
    except Exception:
        return "requester"

def _resolve_sender_smtp(mail_item) -> Optional[str]:
    try:
        s = getattr(mail_item, "SenderEmailAddress", None)
        if s and _simple_email_valid(s):
            return s
    except Exception:
        pass

    try:
        sender = getattr(mail_item, "Sender", None)
        if sender:
            ae = getattr(sender, "AddressEntry", None) or getattr(mail_item, "Sender", None)
            if ae:
                try:
                    ex_user = ae.GetExchangeUser()
                    if ex_user:
                        smtp = getattr(ex_user, "PrimarySmtpAddress", None)
                        if smtp and _simple_email_valid(smtp):
                            return smtp
                except Exception:
                    pass
    except Exception:
        pass

    try:
        ae = getattr(mail_item, "AddressEntry", None)
        if ae:
            try:
                ex_user = ae.GetExchangeUser()
                if ex_user:
                    smtp = getattr(ex_user, "PrimarySmtpAddress", None)
                    if smtp and _simple_email_valid(smtp):
                        return smtp
            except Exception:
                pass
    except Exception:
        pass

    return None

def send_acknowledgement(outlook_ns, original_mail, ticket_obj: dict, ticket_db_obj) -> bool:
    """
    Sends a reply to original_mail acknowledging ticket creation.
    Returns True if send attempted, False on error or if disabled.
    """
    if not REPLY_ENABLED:
        logger.debug("Auto-reply disabled via REPLY_ENABLED")
        return False

    try:
        reporter = _safe_get_reporter_name(original_mail)
        subject = ticket_obj.get("subject") or getattr(original_mail, "Subject", "") or ""
        priority = ticket_obj.get("priority", "medium")

        reporter_email = ticket_obj.get("reporter_email") or _resolve_sender_smtp(original_mail)
        if not reporter_email:
            try:
                raw = getattr(original_mail, "SenderEmailAddress", None)
                if raw and _simple_email_valid(raw):
                    reporter_email = raw
            except Exception:
                reporter_email = None

        ticket_id_str = str(getattr(ticket_db_obj, "id", "") or "")

        reply_subject = REPLY_SUBJECT_TEMPLATE.format(subject=subject, ticket_id=ticket_id_str)
        reply_body = REPLY_BODY_TEMPLATE.format(
            reporter_name_or_email=reporter,
            subject=subject,
            priority=priority,
            ticket_id=ticket_id_str,
            reporter_email=reporter_email or "",
            assignee=getattr(ticket_db_obj, "assigned_to", "") if hasattr(ticket_db_obj, "assigned_to") else ""
        )

        # Expand literal \n sequences from .env into actual newlines
        reply_body = reply_body.replace("\\n", "\n")
        reply_subject = (reply_subject or "").strip()
        reply_body = (reply_body or "").strip()

        logger.debug("Rendered reply_subject: %s", reply_subject[:1000])
        logger.debug("Rendered reply_body (repr): %r", reply_body[:2000])

        # Build simple HTML snippet from plain text
        def _to_html(sn: str) -> str:
            lines = sn.splitlines()
            escaped = [_html.escape(line) for line in lines]
            return "<div>" + "<br/>".join(escaped) + "</div>"

        plain = reply_body
        html_snippet = _to_html(plain)
        validated_ccs = _filter_cc_list(REPLY_CC)

        if not reporter_email or not _simple_email_valid(reporter_email):
            logger.warning("No valid reporter SMTP address found; skipping acknowledgement send. reporter_email=%s", reporter_email)
            return False

        # Primary: try original_mail.Reply()
        try:
            reply = original_mail.Reply()
            existing_html = getattr(reply, "HTMLBody", "") or ""
            existing_plain = getattr(reply, "Body", "") or ""

            if existing_html:
                reply.HTMLBody = html_snippet + "<br/><br/>" + existing_html
            else:
                reply.HTMLBody = html_snippet

            reply.Body = plain + "\n\n" + existing_plain
            reply.Subject = reply_subject

            if validated_ccs:
                reply.CC = ";".join(validated_ccs)

            try:
                reply.Send()
                logger.info("Sent acknowledgement (Reply) for ticket #%s to %s (cc=%s)", ticket_db_obj.id, reporter_email, validated_ccs)
                return True
            except Exception as e:
                logger.warning("Reply.Send() failed; will try fallback CreateItem. error=%s", e)

        except Exception as e:
            logger.warning("Reply() creation failed; falling back to CreateItem. error=%s", e)

        # Fallback: create new mail item and send
        try:
            mail_item = outlook_ns.Application.CreateItem(0)  # olMailItem
            mail_item.To = reporter_email
            mail_item.Subject = reply_subject
            mail_item.Body = plain
            try:
                mail_item.HTMLBody = html_snippet
            except Exception:
                pass
            if validated_ccs:
                mail_item.CC = ";".join(validated_ccs)

            mail_item.Send()
            logger.info("Sent fallback acknowledgement for ticket #%s to %s (cc=%s)", ticket_db_obj.id, reporter_email, validated_ccs)
            return True

        except Exception as e:
            logger.exception("Fallback CreateItem send failed: %s", e)
            return False

    except Exception as e:
        logger.exception("Failed to send acknowledgement for ticket: %s", e)
        return False

# -----------------------
# Ticket creation + move helper
# -----------------------
def _create_ticket_and_move(session, inbox, mail, ticket_obj, sender, entry_id):
    # create Ticket (SQLAlchemy model assumed)
    if get_session is None or Ticket is None:
        logger.error("DB models not available; cannot create ticket.")
        return False

    try:
        t = Ticket(
            subject=ticket_obj.get("subject") or getattr(mail, "Subject", "")[:512],
            description=(ticket_obj.get("description") or getattr(mail, "Body", "") )[:2000],
            reporter_email=ticket_obj.get("reporter_email") or sender,
            priority=ticket_obj.get("priority") or "medium",
            device=ticket_obj.get("device"),
            location=ticket_obj.get("location"),
            tags=json.dumps(ticket_obj.get("tags", [])),
            suggested_actions=json.dumps(ticket_obj.get("suggested_actions", [])),
            confidence=float(ticket_obj.get("confidence", 0.0) or 0.0),
            llm_used=bool(ticket_obj.get("llm_used", False)),
            message_id=entry_id,
            status="open",
        )
        session.add(t)
        session.commit()
        logger.info("Created ticket id=%s for EntryID=%s", t.id, entry_id)

    except Exception as e:
        logger.exception("Ticket creation failed: %s", e)
        session.rollback()
        return False

    # Send acknowledgement (best-effort)
    try:
        outlook_ns = None
        try:
            outlook_ns = mail._Application.GetNamespace("MAPI")
        except Exception:
            try:
                outlook_ns = mail.Application.GetNamespace("MAPI")
            except Exception:
                outlook_ns = None

        if outlook_ns:
            send_acknowledgement(outlook_ns, mail, ticket_obj, t)
        else:
            logger.debug("Outlook namespace not available; skipping acknowledgement send.")
    except Exception as e:
        logger.exception("Acknowledgement sending failed: %s", e)

    # move mail AFTER ticket creation and ack send attempt
    try:
        try:
            processed_folder = inbox.Folders[PROCESSED_FOLDER_NAME]
        except Exception:
            processed_folder = None
        if processed_folder:
            mail.Move(processed_folder)
            logger.info("Moved EntryID=%s to %s", entry_id, PROCESSED_FOLDER_NAME)
    except Exception as e:
        logger.exception("Failed to move mail EntryID=%s: %s", entry_id, e)

    return True

# -----------------------
# Main message processor (heuristic -> LLM escalation)
# -----------------------
def process_single_message_with_escalation(ns, inbox, mail):
    entry_id = getattr(mail, "EntryID", None)
    subject = getattr(mail, "Subject", "") or ""
    body = getattr(mail, "Body", "") or ""
    sender = _extract_sender_email(mail) or getattr(mail, "SenderEmailAddress", None)

    logger.info("Processing EntryID=%s Subject=%r", entry_id, subject)

    if not entry_id:
        logger.warning("Mail has no EntryID → skipping")
        return False

    if get_session is None or Ticket is None:
        logger.error("DB models unavailable; skipping processing")
        return False

    session = get_session()

    try:
        # idempotency: check before heavy work
        exists = session.query(Ticket).filter(Ticket.message_id == entry_id).first()
        if exists:
            logger.info("Skipping (already processed) %s", entry_id)
            return False

        # quick conservative decision whether to process at all
        should_proc, reason = should_process_message(mail)
        if not should_proc:
            logger.info("Skipping message %s: %s", entry_id, reason)
            return False

        # HEURISTIC parse
        heur = {"confidence": 0.0}
        try:
            if heuristic_parser is not None:
                # try modern signature first
                try:
                    heur = heuristic_parser(subject=subject, body=body, headers="")
                except TypeError:
                    # fallback older signature parse(msg)
                    heur = heuristic_parser(mail)
                except Exception as e:
                    logger.exception("Heuristic parser error: %s", e)
                    heur = {"confidence": 0.0}
            else:
                logger.debug("No heuristic parser available")
        except Exception as e:
            logger.exception("Heuristic parser failed: %s", e)
            heur = {"confidence": 0.0}

        heur_conf = float(heur.get("confidence", 0.0) or 0.0)
        logger.info("Heuristic confidence=%.2f", heur_conf)
        if heur_conf >= HEURISTIC_CONF_THRESHOLD:
            heur["llm_used"] = False
            return _create_ticket_and_move(session, inbox, mail, heur, sender, entry_id)

        # If heuristics low confidence, check triggers (subject/body/to/cc/html)
        html_body = ""
        try:
            html_body = getattr(mail, "HTMLBody", "") or ""
        except Exception:
            html_body = ""
        to_field = getattr(mail, "To", "") or ""
        cc_field = getattr(mail, "CC", "") or ""

        trigger_found = _contains_trigger([subject, body, html_body, to_field, cc_field])
        if not trigger_found:
            logger.info("No sysadmin trigger. Sample snippet: %r", (subject + " " + (body or html_body))[:300].replace("\n", " "))
            return False

        # LLM escalation (cached)
        logger.info("Escalating to LLM due to trigger word.")
        cache_key = entry_id
        cached = _llm_cache_get(cache_key)
        if cached:
            parsed = cached.get("parsed")
        else:
            try:
                if parse_email_with_llm is None:
                    logger.error("LLM provider not available; skipping")
                    return False
                parsed = parse_email_with_llm(subject, body, html_body, message_id=entry_id)
                _llm_cache_set(cache_key, {"parsed": parsed})
            except Exception as e:
                logger.exception("LLM error: %s", e)
                return False

        # parsed may be a dict (single) or contain {"multiple_tickets": True, "tickets": [...]}
        tickets = parsed.get("tickets", [parsed]) if isinstance(parsed, dict) else [parsed]

        created_any = False
        for t in tickets:
            conf = float(t.get("confidence", 0.0) or 0.0)
            if conf < LLM_CONF_THRESHOLD:
                logger.info("LLM confidence too low (%.2f). Skipping.", conf)
                continue
            t["llm_used"] = True
            if _create_ticket_and_move(session, inbox, mail, t, sender, entry_id):
                created_any = True

        return created_any

    finally:
        session.close()

# -----------------------
# Poll loop
# -----------------------
def poll_outlook_loop(run_once: bool = False):
    if Dispatch is None:
        raise RuntimeError("pywin32 (win32com) not installed; install with: pip install pywin32")

    ns = _get_outlook_namespace()
    inbox = _get_inbox_folder = None
    # locate inbox (attempt account-specific if provided)
    try:
        if OUTLOOK_ACCOUNT:
            # try to find a root folder matching OUTLOOK_ACCOUNT
            found = None
            for i in range(ns.Folders.Count):
                root = ns.Folders.Item(i + 1)
                if root and (root.Name.lower() == OUTLOOK_ACCOUNT.lower() or getattr(root, "DisplayName", "").lower() == OUTLOOK_ACCOUNT.lower()):
                    try:
                        inbox = root.Folders[INBOX_FOLDER_NAME]
                        found = True
                        break
                    except Exception:
                        pass
            if not inbox:
                inbox = ns.GetDefaultFolder(6)
        else:
            inbox = ns.GetDefaultFolder(6)  # olFolderInbox
    except Exception:
        inbox = ns.GetDefaultFolder(6)

    # ensure processed folder exists under inbox
    try:
        processed_folder = inbox.Folders[PROCESSED_FOLDER_NAME]
    except Exception:
        try:
            processed_folder = inbox.Folders.Add(PROCESSED_FOLDER_NAME)
        except Exception:
            processed_folder = None

    logger.info("Starting Outlook watcher. Poll interval: %s seconds", POLL_INTERVAL)

    try:
        while True:
            try:
                items = inbox.Items.Restrict("[UnRead] = True")
                count = items.Count
                if count:
                    to_check = min(count, MAX_MAILS_PER_POLL)
                    logger.info("Found %d unread messages; checking top %d", count, to_check)
                    # iterate newest-first
                    start = count
                    end = max(count - to_check + 1, 1)
                    for i in range(start, end - 1, -1):
                        try:
                            msg = items.Item(i)
                            # idempotency quick-check
                            session = None
                            try:
                                if get_session is None or Ticket is None:
                                    logger.error("DB models not available; cannot check idempotency")
                                    continue
                                session = get_session()
                                entry_id = getattr(msg, "EntryID", None)
                                if session.query(Ticket).filter(Ticket.message_id == entry_id).first():
                                    logger.info("Message %s already processed, skipping", entry_id)
                                    continue
                            except Exception:
                                logger.exception("Idempotency check failed")
                            finally:
                                if session:
                                    session.close()

                            # process message
                            process_single_message_with_escalation(ns, inbox, msg)
                        except Exception:
                            logger.exception("Failed to handle item at index %s", i)
                else:
                    logger.debug("No unread messages")
            except Exception:
                logger.exception("Polling loop error")

            if run_once:
                logger.info("run_once flag set — exiting after one poll iteration")
                break

            time.sleep(POLL_INTERVAL)
    except KeyboardInterrupt:
        logger.info("Watcher stopped by user")

# -----------------------
# CLI / entrypoint
# -----------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Outlook watcher")
    parser.add_argument("--once", action="store_true", help="Run one poll iteration and exit")
    args = parser.parse_args()

    env_once = os.getenv("WATCHER_RUN_ONCE", "").strip().lower() in ("1", "true", "yes")
    run_once_flag = args.once or env_once

    try:
        poll_outlook_loop(run_once=run_once_flag)
    except Exception:
        logger.exception("Unhandled exception in watcher main")
        raise
