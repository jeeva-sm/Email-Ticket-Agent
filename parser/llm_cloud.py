# parser/llm_cloud.py
import os
import time
import json
import hashlib
import logging
import re
from typing import Any, Dict, Optional, List
from pydantic import BaseModel, Field, ValidationError, conlist, confloat

from parser.heuristic import heuristic_fallback
from parser.openai_client import call_chat_completion, OPENAI_MODEL, OPENAI_MAX_TOKENS

logger = logging.getLogger(__name__)

CACHE_DIR = os.path.join(os.path.dirname(__file__), "..", ".cache_llm")
os.makedirs(CACHE_DIR, exist_ok=True)
MAX_BODY_CHARS = int(os.getenv("LLM_MAX_BODY_CHARS", "1200"))
LLM_MAX_TOKENS = int(os.getenv("OPENAI_MAX_TOKENS", str(OPENAI_MAX_TOKENS if 'OPENAI_MAX_TOKENS' in globals() else 512)))

class TicketOut(BaseModel):
    subject: str
    description: str
    reporter_email: Optional[str] = None
    priority: str = Field(..., regex="^(low|medium|high|critical)$")
    device: Optional[str] = None
    location: Optional[str] = None
    tags: conlist(str, min_items=0) = []
    suggested_actions: conlist(str, min_items=0) = []
    confidence: confloat(ge=0.0, le=1.0)


def _build_prompt(subject: str, body: str, headers: str) -> str:
    body = (body or "").strip()
    if len(body) > MAX_BODY_CHARS:
        body = body[-MAX_BODY_CHARS:]
    prompt = (
        f"Subject: {subject}\nHeaders:\n{headers}\n\nBody:\n{body}\n\n"
        "Extract the issue(s) from this email and return a JSON array of ticket objects exactly matching the schema.\n"
        "Do NOT add any other text, explanation, or markdown. The JSON must be valid and parseable."
    )
    return prompt


def _cache_path(key: str) -> str:
    h = hashlib.sha256(key.encode()).hexdigest()
    return os.path.join(CACHE_DIR, h + ".json")


def _cache_get(key: str) -> Optional[Dict[str, Any]]:
    p = _cache_path(key)
    if not os.path.exists(p):
        return None
    try:
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _cache_set(key: str, obj: Dict[str, Any]):
    p = _cache_path(key)
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(obj, f)
    except Exception as e:
        logger.warning("Failed writing cache %s: %s", p, e)


def _extract_json(text: str):
    text = (text or "").strip()
    if not text:
        raise ValueError("Empty LLM output")

    # 1) direct JSON
    try:
        return json.loads(text)
    except Exception:
        pass

    # 2) fences
    m = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text, re.IGNORECASE)
    if m:
        candidate = m.group(1).strip()
        try:
            return json.loads(candidate)
        except Exception:
            text = candidate

    # 3) array/object extraction
    m = re.search(r"(\[[\s\S]*\])", text)
    if m:
        return json.loads(m.group(1))
    m = re.search(r"(\{[\s\S]*\})", text)
    if m:
        return json.loads(m.group(1))

    # 4) recover truncated arrays by attempting to find a balanced substring
    start = text.find('[')
    if start != -1:
        depth = 0
        for i in range(start, len(text)):
            if text[i] == '[':
                depth += 1
            elif text[i] == ']':
                depth -= 1
                if depth == 0:
                    candidate = text[start:i+1]
                    try:
                        return json.loads(candidate)
                    except Exception:
                        break
        candidate = text[start:].rstrip()
        try:
            return json.loads(candidate + "]")
        except Exception:
            pass

    # 5) try single object repair
    start = text.find('{')
    if start != -1:
        candidate = text[start:].rstrip()
        try:
            return json.loads(candidate + "}")
        except Exception:
            pass

    raise ValueError("No JSON found in LLM output")


def _cleanup_to_json(raw_text: str) -> Optional[str]:
    """
    Call the LLM with a short, deterministic prompt that asks it to extract/repair the JSON
    from the given raw_text and RETURN the exact schema including required fields.
    """
    if not raw_text:
        return None

    system = (
        "You are a strict JSON repair assistant. The user will provide some raw model output. "
        "Return ONLY valid JSON (array or object) that follows this schema EXACTLY:\n"
        "- subject (string)\n"
        "- description (string)\n"
        "- reporter_email (string|null)\n"
        "- priority (one of: low, medium, high, critical)\n"
        "- device (string|null)\n"
        "- location (string|null)\n"
        "- tags (array)\n"
        "- suggested_actions (array)\n"
        "- confidence (number between 0.0 and 1.0)\n\n"
        "Do NOT include any other keys or explanation. If you cannot determine a field, set it to null (except confidence: set 0.5)."
    )

    user = (
        "RAW OUTPUT (from the first model call):\n\n" + raw_text + "\n\n"
        "Return the repaired JSON now."
    )

    try:
        cleaned = call_chat_completion(
            messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
            model=OPENAI_MODEL,
            max_tokens=LLM_MAX_TOKENS,
        )
        cleaned = (cleaned or "").strip()
        # Extract JSON if there's surrounding text
        m = re.search(r"(\[[\s\S]*\])", cleaned)
        if m:
            return m.group(1)
        m = re.search(r"(\{[\s\S]*\})", cleaned)
        if m:
            return m.group(1)
        # If returned text already seems JSON-like starting with [ or {
        if cleaned.startswith('[') or cleaned.startswith('{'):
            return cleaned
        return None
    except Exception as e:
        logger.exception("Cleanup LLM call failed: %s", e)
        return None



def parse_email_with_openai(subject: str, body: str, headers: str, message_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Call the cloud LLM to extract ticket(s). On success returns a ticket dict or multiple-tickets dict.
    Falls back to heuristic_fallback(...) with confidence=0.0 and llm_used=False on failures.
    This version performs a cleanup LLM pass if the initial output is not valid JSON,
    and fills missing required fields (priority, confidence, lists) before validation.
    """
    prompt = _build_prompt(subject, body, headers)
    cache_key = message_id or hashlib.sha256(prompt.encode()).hexdigest()

    raw = None
    cached = _cache_get(cache_key)
    if cached and isinstance(cached, dict) and "raw" in cached:
        raw = cached["raw"]
        logger.info("LLM cache hit for key=%s", cache_key)

    # If no cached raw, call the LLM
    if not raw:
        system_msg = (
            "You are a machine-readable extractor. YOU MUST reply with ONLY valid JSON and nothing else.\n"
            "Return an ARRAY of objects. Each object MUST contain these keys:\n"
            "  subject, description, reporter_email|null, priority (low|medium|high|critical),\n"
            "  device|null, location|null, tags[], suggested_actions[], confidence (0.0-1.0).\n"
            "Return compact JSON (no extra whitespace)."
        )
        messages = [
            {"role": "system", "content": system_msg},
            {"role": "user", "content": prompt},
        ]

        try:
            logger.info("Calling cloud LLM model=%s prompt_len=%d", OPENAI_MODEL, len(prompt))
            raw = call_chat_completion(messages=messages, model=OPENAI_MODEL, max_tokens=LLM_MAX_TOKENS)
            _cache_set(cache_key, {"raw": raw, "ts": int(time.time())})
        except Exception as e:
            logger.exception("Cloud LLM call failed: %s", e)
            fallback = heuristic_fallback(subject, body, headers)
            fallback["confidence"] = 0.0
            fallback["llm_used"] = False
            return fallback

    # Try extracting JSON directly from raw output
    payload = None
    try:
        payload = _extract_json(raw)
    except Exception as e:
        logger.warning("Initial extraction failed: %s. Attempting cleanup step.", e)
        # Try cleanup/repair pass via LLM
        try:
            cleaned = _cleanup_to_json(raw)
        except Exception as e2:
            logger.exception("Cleanup attempt raised: %s", e2)
            cleaned = None

        if cleaned:
            try:
                payload = _extract_json(cleaned)
            except Exception as e3:
                logger.exception("Extraction from cleaned text failed: %s", e3)
                payload = None
        else:
            payload = None

    # If still no payload, save raw for inspection and fallback
    if not payload:
        debug_path = _cache_path(cache_key) + ".bad.txt"
        try:
            with open(debug_path, "w", encoding="utf-8") as f:
                f.write(raw or "")
            logger.warning("Saved bad LLM output -> %s", debug_path)
        except Exception:
            logger.exception("Failed to save bad LLM output for inspection")
        fallback = heuristic_fallback(subject, body, headers)
        fallback["confidence"] = 0.0
        fallback["llm_used"] = False
        return fallback

    # Normalize single object -> list
    if isinstance(payload, dict):
        payload = [payload]

    validated: List[Dict[str, Any]] = []
    for item in payload:
        # ---- Fill missing keys with sensible defaults BEFORE validation ----
        # priority: infer via heuristic if missing/empty
        if "priority" not in item or not item.get("priority"):
            try:
                hf = heuristic_fallback(subject, body, headers)
                item["priority"] = hf.get("priority", "medium")
            except Exception:
                item["priority"] = "medium"

        # confidence: if missing or None, give a conservative LLM confidence
        if "confidence" not in item or item.get("confidence") is None:
            # prefer a slightly optimistic default but not too high
            item["confidence"] = 0.6

        # tags and suggested_actions should be lists
        if "tags" not in item or not isinstance(item.get("tags"), list):
            item["tags"] = item.get("tags") or []
        if "suggested_actions" not in item or not isinstance(item.get("suggested_actions"), list):
            item["suggested_actions"] = item.get("suggested_actions") or []

        # ensure reporter_email key exists (can be null)
        if "reporter_email" not in item:
            item["reporter_email"] = None

        # device/location: ensure keys exist (can be null)
        if "device" not in item:
            item["device"] = None
        if "location" not in item:
            item["location"] = None

        # ---- Now validate with Pydantic ----
        try:
            t = TicketOut(**item)
            validated.append(t.dict())
        except ValidationError as e:
            logger.warning("Validation failed for item after fill-ins: %s error=%s", str(item)[:400], e)
            # For safety, fallback to heuristic when validation still fails
            fallback = heuristic_fallback(subject, body, headers)
            fallback["confidence"] = 0.0
            fallback["llm_used"] = False
            return fallback

    # Return a single ticket or multiple tickets wrapper
    if len(validated) == 1:
        validated[0]["llm_used"] = True
        return validated[0]

    return {"multiple_tickets": True, "tickets": validated, "llm_used": True}
