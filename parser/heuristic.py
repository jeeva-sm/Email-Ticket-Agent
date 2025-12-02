# parser/heuristic.py
"""
Heuristic email-to-ticket parser.

This module provides parse_email_message(...) which tries to extract
ticket fields using deterministic rules and produces a confidence score.

It accepts two call styles:
  - parse_email_message(subject, body, headers)
  - parse_email_message(mailitem)  # Outlook MailItem object

Returns a dict with keys:
  subject, description, reporter_email, priority, device, location,
  tags (list), suggested_actions (list), confidence (0.0-1.0)
"""
import re
import logging
from typing import Optional, Dict, Any, Tuple, Union, List

logger = logging.getLogger(__name__)

# Primary indicators (increase confidence)
SYSADMIN_GREETINGS = [
    "hi sysadmin", "dear sysadmin", "dear sys admin", "hi sys admin",
    "to sysadmin", "sysadmin team", "sys admin team", "@sysadmin", "hi team sysadmin"
]
IT_KEYWORDS = [
    "internet", "network", "ethernet", "wifi", "wi-fi", "lan", "wan",
    "anydesk", "vpn", "openvpn", "printer", "printer", "access", "login",
    "password", "connect", "connection", "install", "remote", "anydesk id",
    "anydeskid", "anydesk id:", "anydesk id", "id:", "ip address", "dhcp"
]
URGENCY_WORDS = ["asap", "urgent", "immediately", "important", "need it", "please help", "soon"]
ACTION_VERBS = ["install", "connect", "configure", "reset", "update", "patch", "access", "login"]
LOCATION_PATTERNS = [
    r"bay location\s*[:\-]?\s*([A-Za-z0-9\-\s]+)",
    r"bay\s*[:\-]?\s*([A-Za-z0-9\-\s]+)",
    r"location\s*[:\-]?\s*([A-Za-z0-9\-\s]+)",
    r"floor\s*[:\-]?\s*([A-Za-z0-9\-\s]+)"
]

# weighting (tunable)
WEIGHT_GREETING = 0.15
WEIGHT_IT_KEYWORD = 0.25
WEIGHT_URGENCY = 0.15
WEIGHT_ACTION_VERB = 0.10
WEIGHT_LOCATION = 0.15
WEIGHT_DEVICE_EVIDENCE = 0.2  # e.g., explicit "Ethernet cable", "AnyDesk ID"

# priority mapping heuristics
PRIORITY_KEYWORDS = {
    "critical": ["server down", "production down", "service down"],
    "high": ["asap", "urgent", "immediately", "can't connect", "cannot connect", "no internet", "network down"],
    "medium": ["not working", "issue", "problem", "request"],
    "low": ["info", "request for info", "how to"]
}

def _normalize_text(s: Optional[str]) -> str:
    if not s:
        return ""
    return re.sub(r"\s+", " ", s.strip()).lower()

def _extract_location(text: str) -> Optional[str]:
    for pat in LOCATION_PATTERNS:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            loc = m.group(1).strip()
            return loc
    return None

def _detect_device(text: str) -> Optional[str]:
    # simple canonical device detection
    t = text.lower()
    if "anydesk" in t:
        return "anydesk"
    if "openvpn" in t or "vpn" in t:
        return "vpn"
    if "ethernet" in t or "ethernet cable" in t or "lan" in t:
        return "network"
    if "printer" in t:
        return "printer"
    if "wifi" in t or "wi-fi" in t:
        return "wifi"
    return None

def _choose_priority(subject: str, body: str) -> str:
    txt = f"{subject}\n{body}".lower()
    for p, kws in PRIORITY_KEYWORDS.items():
        for k in kws:
            if k in txt:
                return p
    # default fallback
    # if urgent words present -> high, else medium
    if any(w in txt for w in URGENCY_WORDS):
        return "high"
    return "medium"

def _score_from_evidence(text: str, subject_only: str = "") -> Tuple[float, List[str]]:
    """
    Return (score, evidence_list)
    evidence_list contains short notes about matched signals.
    """
    text_l = text.lower()
    evidence = []
    score = 0.0

    # greeting
    if any(g in text_l for g in SYSADMIN_GREETINGS):
        score += WEIGHT_GREETING
        evidence.append("sysadmin greeting")

    # IT keywords (count distinct)
    found_it = set()
    for kw in IT_KEYWORDS:
        if kw in text_l:
            found_it.add(kw)
    if found_it:
        # larger boost for multiple distinct keywords
        score += WEIGHT_IT_KEYWORD
        if len(found_it) >= 2:
            score += min(0.15, 0.05 * (len(found_it)-1))  # small bonus for multiple hits
        evidence.append(f"it_keywords({','.join(sorted(found_it))})")

    # urgency words
    found_urg = [w for w in URGENCY_WORDS if w in text_l]
    if found_urg:
        score += WEIGHT_URGENCY
        evidence.append(f"urgency({','.join(found_urg)})")

    # action verbs
    found_actions = [v for v in ACTION_VERBS if v in text_l]
    if found_actions:
        score += WEIGHT_ACTION_VERB
        evidence.append(f"action({','.join(found_actions)})")

    # device evidence
    dev = _detect_device(text)
    if dev:
        score += WEIGHT_DEVICE_EVIDENCE
        evidence.append(f"device_detected({dev})")

    # location evidence
    loc = _extract_location(text)
    if loc:
        score += WEIGHT_LOCATION
        evidence.append(f"location({loc})")

    # small floor: if subject explicitly contains IT keyword, give a small boost
    if subject_only and any(k in subject_only for k in IT_KEYWORDS):
        score += 0.05
        evidence.append("subject_it_hint")

    # bound
    if score > 1.0:
        score = 1.0

    return score, evidence

def parse_email_message(*args, **kwargs) -> Dict[str, Any]:
    """
    Flexible wrapper. Accepts either:
       parse_email_message(subject, body, headers)
    or parse_email_message(mailitem)

    Returns a dict: subject, description, reporter_email, priority,
    device, location, tags(list), suggested_actions(list), confidence (float)
    """
    # signature flexibility
    subject = ""
    body = ""
    headers = ""
    reporter_email = None

    if len(args) == 1 and not kwargs:
        # likely called with a MailItem object
        mail = args[0]
        try:
            subject = getattr(mail, "Subject", "") or ""
            # body fallbacks: plain text or HTML stripped
            body = getattr(mail, "Body", "") or ""
            # try sender
            reporter_email = getattr(mail, "SenderEmailAddress", None)
            headers = ""
        except Exception:
            # fallback to empty
            subject = subject or ""
            body = body or ""
    else:
        # expecting (subject, body, headers) maybe as positional
        if len(args) >= 1:
            subject = args[0] or ""
        if len(args) >= 2:
            body = args[1] or ""
        if len(args) >= 3:
            headers = args[2] or ""
        # kwargs overrides
        subject = kwargs.get("subject", subject)
        body = kwargs.get("body", body)
        headers = kwargs.get("headers", headers)
        reporter_email = kwargs.get("reporter_email", reporter_email)

    # normalize
    subj_norm = _normalize_text(subject)
    body_norm = _normalize_text(body)
    combined = f"{subj_norm}\n{body_norm}\n{(headers or '').lower()}"

    # scoring
    score, evidence = _score_from_evidence(combined, subject_only=subj_norm)

    # device/location detection
    device = _detect_device(combined)
    location = _extract_location(combined)

    # suggested actions (simple heuristics)
    suggestions = []
    if device == "anydesk":
        suggestions.append("Request remote connection via AnyDesk")
    if device == "vpn":
        suggestions.append("Verify VPN installation/config and provide instructions")
    if device == "network" or "internet" in combined:
        suggestions.append("Check network link, NIC, and switch port; ask user to test with ping")

    # priority
    priority = _choose_priority(subject, body)

    # produce result
    result = {
        "subject": subject.strip() or "(no subject)",
        "description": (body.strip()[:2000]) if body else "",
        "reporter_email": reporter_email,
        "priority": priority,
        "device": device,
        "location": location,
        "tags": [],  # heuristics can append tags if needed
        "suggested_actions": suggestions,
        "confidence": round(float(score), 3),
    }

    # debug logging for tuning
    logger.debug("Heuristic parse result: subj=%r, score=%s, evidence=%s", subject, result["confidence"], evidence)

    return result
# --- compatibility shim for older callers -------------------------------------------------
def heuristic_fallback(subject=None, body=None, headers=None, mailitem=None):
    """
    Backwards-compatible wrapper used by llm_cloud and other modules.
    Accepts either:
      - heuristic_fallback(subject, body, headers)
      - heuristic_fallback(mailitem=outlook_mailitem)
    Returns the same dict shape as parse_email_message().
    """
    # Support both calling styles
    try:
        if mailitem is not None:
            # callers that pass the MailItem by keyword
            return parse_email_message(mailitem)
        # callers that pass (subject, body, headers)
        return parse_email_message(subject or "", body or "", headers or "")
    except TypeError:
        # fallback in case some callers use different signature
        try:
            # try plain parse call with whatever was given
            return parse_email_message(subject, body, headers)
        except Exception as e:
            # ultimate fallback: return a minimal dict so callers can proceed
            logger.exception("heuristic_fallback failed: %s", e)
            return {
                "subject": subject or "",
                "description": body or "",
                "reporter_email": None,
                "priority": "medium",
                "device": None,
                "location": None,
                "tags": [],
                "suggested_actions": [],
                "confidence": 0.0,
            }
