# parser/llm_provider.py
import os
import logging
from typing import Any, Dict, Optional

logger = logging.getLogger(__name__)

PROVIDER = os.getenv("LLM_PROVIDER", "cloud").lower()

# Default imports are lazy to avoid importing heavy libs at module import time.


def _cloud_provider():
    # cloud provider (OpenAI) or OpenAI-compatible local servers (Ollama)
    try:
        from parser.llm_cloud import parse_email_with_openai
        return parse_email_with_openai
    except Exception as e:
        logger.exception("Failed to import cloud LLM parser: %s", e)
        raise


def _lmstudio_provider():
    try:
        # local LM Studio parser (already exists in your repo)
        from parser.llm_parser import parse_email_with_llm as lmstudio_parse
        return lmstudio_parse
    except Exception as e:
        logger.exception("Failed to import LM Studio parser: %s", e)
        raise


def _llama_local_provider():
    """
    If you implement llama-cpp-python wrapper 'parser/llama_local.py' with
    call_local(prompt, max_tokens=...) returning a raw string, we wrap it into
    a parse-like function (simplified). You can enhance this wrapper later.
    """
    try:
        from parser.llama_local import call_local as llama_call
    except Exception as e:
        logger.exception("Failed to import llama_local: %s", e)
        raise

    def _wrapper(subject: str, body: str, headers: str, message_id: Optional[str] = None) -> Dict[str, Any]:
        prompt = f"Subject: {subject}\n\nHeaders:\n{headers}\n\nBody:\n{body}"
        raw = llama_call(prompt, max_tokens=int(os.getenv("LLM_MAX_TOKENS", "120")))
        # Minimal best-effort return: put raw into description and low confidence
        return {
            "subject": subject or "No subject",
            "description": raw if isinstance(raw, str) else str(raw),
            "reporter_email": None,
            "priority": "medium",
            "device": None,
            "location": None,
            "tags": [],
            "suggested_actions": [],
            "confidence": 0.5,
            "llm_used": True,
            "llm_raw": raw,
        }

    return _wrapper


# Dispatcher factory
if PROVIDER in ("cloud", "ollama_local"):
    parse_email_with_llm = _cloud_provider()
elif PROVIDER == "lmstudio":
    parse_email_with_llm = _lmstudio_provider()
elif PROVIDER == "llama_local":
    parse_email_with_llm = _llama_local_provider()
else:
    raise RuntimeError(f"Unknown LLM_PROVIDER '{PROVIDER}'. Set LLM_PROVIDER env to one of: cloud, ollama_local, lmstudio, llama_local")
