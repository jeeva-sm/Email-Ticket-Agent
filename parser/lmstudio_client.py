# parser/lmstudio_client.py
import os
import requests
import logging
import json

logger = logging.getLogger(__name__)

LMSTUDIO_URL = os.getenv("LMSTUDIO_URL", "http://127.0.0.1:1234")
LMSTUDIO_MODEL = os.getenv("LMSTUDIO_MODEL", "mistralai/mistral-7b-instruct-v0.3")

def _post(url: str, payload: dict, timeout: int = 30):
    try:
        r = requests.post(url, json=payload, timeout=timeout)
        return r
    except Exception as e:
        logger.exception("HTTP POST failed: %s %s", url, e)
        raise

def call_lmstudio(prompt: str, max_tokens: int = 800, temperature: float = 0.0) -> str:
    """
    Call LM Studio local API. Prefer /v1/completions with prompt for instruct models.
    Fallback to /v1/chat/completions using only 'user' role if needed.
    """
    base = LMSTUDIO_URL.rstrip("/")
    model = LMSTUDIO_MODEL or None

    # 1) Try completions endpoint (instruct-style)
    try:
        url = base + "/v1/completions"
        payload = {
            "model": model,
            "prompt": prompt,
            "max_tokens": max_tokens,
            "temperature": temperature,
        }
        r = _post(url, payload)
        if r.status_code == 200:
            # Try to return text from standard shapes
            try:
                data = r.json()
                # Common shapes: data["choices"][0]["text"]
                if "choices" in data and data["choices"]:
                    ch = data["choices"][0]
                    if isinstance(ch, dict) and "text" in ch:
                        return ch["text"]
                # fallback to stringified JSON
                return json.dumps(data)
            except Exception:
                return r.text
        else:
            logger.debug("Completions endpoint returned status %s: %s", r.status_code, r.text[:400])
    except Exception as e:
        logger.debug("Completions call failed: %s", e)

    # 2) Fallback: chat completions but use only "user" role (model complains about system role)
    try:
        url = base + "/v1/chat/completions"
        payload = {
            "model": model,
            "messages": [
                # use only user role; do not send system role to avoid the jinja-template error
                {"role": "user", "content": prompt}
            ],
            "temperature": temperature,
            "max_tokens": max_tokens,
        }
        r = _post(url, payload)
        if r.status_code == 200:
            try:
                data = r.json()
                # standard chat completion shape
                if "choices" in data and data["choices"]:
                    msg = data["choices"][0].get("message", {})
                    # message may contain "content" or "role"/"content"
                    if isinstance(msg, dict) and "content" in msg:
                        return msg["content"]
                return json.dumps(data)
            except Exception:
                return r.text
        else:
            logger.debug("Chat endpoint returned status %s: %s", r.status_code, r.text[:400])
            # raise to inform caller
            raise RuntimeError(f"LM Studio chat endpoint error: {r.status_code} {r.text[:400]}")
    except Exception as e:
        logger.exception("LM Studio fallback chat call failed: %s", e)
        raise RuntimeError("LM Studio calls failed; check LMStudio server logs and model availability.") from e
