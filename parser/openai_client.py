# parser/openai_client.py
import load_env
import os
import time
import logging
from typing import Optional, List, Dict, Any

import requests

# Optional: import OpenAI SDK only when needed
try:
    from openai import OpenAI
except Exception:
    OpenAI = None  # meaningful check later

logger = logging.getLogger(__name__)

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL", "").rstrip("/")  # e.g. http://127.0.0.1:11434/v1
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
OPENAI_TIMEOUT = int(os.getenv("OPENAI_TIMEOUT", "120"))
OPENAI_MAX_TOKENS = int(os.getenv("OPENAI_MAX_TOKENS", "150"))
OPENAI_TEMPERATURE = float(os.getenv("OPENAI_TEMPERATURE", "0.0"))


# ------------------------
# Helpers
# ------------------------
def _log_call_start(model: str, provider: str):
    logger.info("LLM call provider=%s model=%s", provider, model)


# ------------------------
# Requests-based path (for Ollama / custom base URL)
# ------------------------
def _call_via_requests_chat(messages: List[Dict[str, str]], model: str, max_tokens: int, temperature: float, timeout: int) -> str:
    """
    Calls OPENAI_BASE_URL/v1/chat/completions using requests.
    Ollama accepts OpenAI-compatible chat/completions path.
    """
    url = OPENAI_BASE_URL + "/v1/chat/completions"
    payload = {
        "model": model,
        "messages": messages,
        "max_tokens": max_tokens,
        "temperature": temperature,
    }
    headers = {
        "Content-Type": "application/json",
        # Ollama doesn't validate the key but OpenAI-compatible servers expect an Authorization header
        "Authorization": f"Bearer {OPENAI_API_KEY}" if OPENAI_API_KEY else ""
    }
    r = requests.post(url, json=payload, headers=headers, timeout=timeout)
    r.raise_for_status()
    data = r.json()
    # Common shapes: choices[0].message.content or choices[0].text
    try:
        return data["choices"][0]["message"]["content"]
    except Exception:
        try:
            return data["choices"][0]["text"]
        except Exception:
            return str(data)


def _call_via_requests_completions(prompt: str, model: str, max_tokens: int, temperature: float, timeout: int) -> str:
    url = OPENAI_BASE_URL + "/v1/completions"
    payload = {"model": model, "prompt": prompt, "max_tokens": max_tokens, "temperature": temperature}
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {OPENAI_API_KEY}" if OPENAI_API_KEY else ""}
    r = requests.post(url, json=payload, headers=headers, timeout=timeout)
    r.raise_for_status()
    data = r.json()
    try:
        return data["choices"][0]["text"]
    except Exception:
        return str(data)


# ------------------------
# SDK-based path (for real OpenAI)
# ------------------------
def _make_sdk_client() -> Optional[OpenAI]:
    if OpenAI is None:
        return None
    # rely on SDK to pick up OPENAI_API_KEY from env or passed value
    if OPENAI_API_KEY:
        return OpenAI(api_key=OPENAI_API_KEY)
    else:
        return OpenAI()  # will rely on env var OPENAI_API_KEY


def _call_via_sdk_chat(messages: List[Dict[str, str]], model: str, max_tokens: int, temperature: float, timeout: int) -> str:
    client = _make_sdk_client()
    if client is None:
        raise RuntimeError("OpenAI SDK not available in environment")
    resp = client.chat.completions.create(model=model, messages=messages, max_tokens=max_tokens, temperature=temperature, timeout=timeout)
    try:
        return resp.choices[0].message.content
    except Exception:
        try:
            return resp["choices"][0]["message"]["content"]
        except Exception:
            return str(resp)


def _call_via_sdk_completions(prompt: str, model: str, max_tokens: int, temperature: float, timeout: int) -> str:
    client = _make_sdk_client()
    if client is None:
        raise RuntimeError("OpenAI SDK not available in environment")
    resp = client.completions.create(model=model, prompt=prompt, max_tokens=max_tokens, temperature=temperature, timeout=timeout)
    try:
        return resp.choices[0].text
    except Exception:
        try:
            return resp["choices"][0]["text"]
        except Exception:
            return str(resp)


# ------------------------
# Public unified function
# ------------------------
def call_chat_completion(messages: List[Dict[str, str]], model: Optional[str] = None,
                         max_tokens: Optional[int] = None, temperature: Optional[float] = None,
                         timeout: Optional[int] = None) -> str:
    model = model or OPENAI_MODEL
    max_tokens = max_tokens or OPENAI_MAX_TOKENS
    temperature = temperature if temperature is not None else OPENAI_TEMPERATURE
    timeout = timeout or OPENAI_TIMEOUT

    provider = OPENAI_BASE_URL or "openai"
    _log_call_start(model, provider)
    start = time.time()

    # If custom base URL set (Ollama), use requests (more reliable)
    if OPENAI_BASE_URL:
        # try chat endpoint first
        try:
            out = _call_via_requests_chat(messages, model, max_tokens, temperature, timeout)
            logger.info("LLM call done provider=%s latency=%.2fs", provider, time.time() - start)
            return out
        except requests.HTTPError as e:
            # bubble up non-auth errors; if 401/403 from a remote OpenAI endpoint, raise
            logger.debug("requests chat failed: %s", e)
            # try completions fallback
            try:
                # convert messages -> prompt
                prompt = "\n".join([f"{m.get('role','user')}: {m.get('content','')}" for m in messages])
                out = _call_via_requests_completions(prompt, model, max_tokens, temperature, timeout)
                logger.info("LLM completions fallback done provider=%s latency=%.2fs", provider, time.time() - start)
                return out
            except Exception as e2:
                logger.exception("Both requests chat and completions failed: %s", e2)
                raise
        except Exception as e:
            logger.exception("Requests-based LLM call failed: %s", e)
            raise

    # Else, use SDK for real OpenAI
    try:
        out = _call_via_sdk_chat(messages, model, max_tokens, temperature, timeout)
        logger.info("LLM SDK call done provider=%s latency=%.2fs", provider, time.time() - start)
        return out
    except Exception as e:
        logger.debug("SDK chat failed: %s", e)
        # try SDK completions as fallback
        try:
            prompt = "\n".join([f"{m.get('role','user')}: {m.get('content','')}" for m in messages])
            out = _call_via_sdk_completions(prompt, model, max_tokens, temperature, timeout)
            logger.info("LLM SDK completions fallback done provider=%s latency=%.2fs", provider, time.time() - start)
            return out
        except Exception as e2:
            logger.exception("Both SDK chat and completions failed: %s", e2)
            raise


    # small helper (put in parser/openai_client.py or a telemetry module)
import time, math, logging
logger = logging.getLogger("llm_telemetry")

def call_with_telemetry(messages, **kwargs):
    from parser.openai_client import call_chat_completion, OPENAI_MODEL
    start = time.time()
    out = call_chat_completion(messages=messages, **kwargs)
    latency = time.time() - start
    # crude token estimate
    prompt_chars = sum(len(m.get("content","")) for m in messages)
    inp_tokens = math.ceil(prompt_chars / 4)
    logger.info("LLM call model=%s latency=%.2fs est_inp_tokens=%d", OPENAI_MODEL, latency, inp_tokens)
    return out

