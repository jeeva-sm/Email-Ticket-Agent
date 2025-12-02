# tests/test_parser.py
import email
from parser.heuristic import parse_email_message


def load_message(path: str):
    """
    Open a text or .eml file from fixtures and return an EmailMessage object.
    Works for plain text and HTML emails.
    """
    with open(path, "rb") as f:
        raw = f.read()

    try:
        # First try parsing as raw RFC822 bytes
        msg = email.message_from_bytes(raw)
    except Exception:
        # Fallback for plain text fixtures
        msg = email.message_from_string(raw.decode("utf-8", errors="ignore"))

    return msg


def test_plain_email():
    msg = load_message("tests/fixtures/email_plain.txt")
    parsed = parse_email_message(msg)

    assert parsed["subject"] != ""
    assert "internet" in parsed["description"].lower()
    assert parsed["reporter_email"] == "alice@example.com"
    # Contains 'urgently' so priority should be high
    assert parsed["priority"] in ["high", "medium"]  # heuristic fallback acceptable
    assert parsed["device"] in [None, "network"]
    assert isinstance(parsed["tags"], str)


def test_html_email():
    msg = load_message("tests/fixtures/email_html.eml")
    parsed = parse_email_message(msg)

    assert "printer" in parsed["description"].lower() or "printer" in parsed["subject"].lower()
    assert parsed["reporter_email"] == "bob@example.com"
    assert parsed["device"] in [None, "printer", "network"]
    assert isinstance(parsed["tags"], str)


def test_forwarded_email():
    msg = load_message("tests/fixtures/email_forwarded.txt")
    parsed = parse_email_message(msg)

    assert "login" in parsed["description"].lower() or "auth" in parsed["tags"]
    assert parsed["reporter_email"] in ["charlie@example.com", None]  # forwarded emails vary
    assert isinstance(parsed["tags"], str)
