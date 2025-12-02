# parser/llm_parser.py

import json
import logging
import re
from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field, ValidationError, conlist, confloat

from parser.heuristic import parse_email_message as heuristic_parser
from parser.lmstudio_client import call_lmstudio  # we will create this next

logger = logging.getLogger(__name__)

# -------------------------
# Pydantic schema for validated LLM output
# -------------------------

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

# -------------------------
# Prompt builder
# -------------------------

def build_prompt(subject: str, body: str, headers: str) -> str:
    """
    Build a structured prompt for Mistral 7B Instruct via LM Studio.
    Output must be a JSON array of ticket objects.
    """
    return f"""
You are an IT helpdesk email parser.

Given the email subject, headers, and body:
- Extract the ticket information
- Return ONLY a JSON array
- Each item in the array must contain EXACTLY these fields:

subject
description
reporter_email
priority  (low|medium|high|critical)
device
location
tags
suggested_actions
confidence  (0.0 - 1.0)

If the email describes only one issue, return an array with one object.

Subject: {subject}

Headers:
{headers}

Body:
{body}

Return ONLY the JSON array. No explanation text.
    """.strip()


# -------------------------
# Extract JSON from model output
# -------------------------

def extract_json(text: str) -> Any:
    """
    Extract a JSON list/object from an LLM output.
    """
    text = text.strip()

    # 1. Try direct parsing
    try:
        return json.loads(text)
    except Exception:
        pass

    # 2. Try to find a JSON array
    m = re.search(r"(\[[\s\S]*\])", text)
    if m:
        return json.loads(m.group(1))

    # 3. Try to find a JSON object
    m = re.search(r"(\{[\s\S]*\})", text)
    if m:
        return json.loads(m.group(1))

    raise ValueError("No valid JSON found in LLM output")


# -------------------------
# Main LLM parser
# -------------------------

def parse_email_with_llm(subject: str, body: str, headers: str) -> Dict[str, Any]:
    """
    High-level LLM parser:
    - Build prompt
    - Call LM Studio
    - Extract JSON
    - Validate JSON structure with Pydantic
    - Return dict or fallback to heuristic
    """
    prompt = build_prompt(subject, body, headers)

    try:
        raw = call_lmstudio(prompt)
        logger.info("LLM raw output: %s", raw[:200])
    except Exception as e:
        logger.error("LM Studio call failed: %s", e)
        fallback = heuristic_parser(subject, body, headers)
        fallback["confidence"] = 0.0
        fallback["llm_used"] = False
        return fallback

    # Try to parse JSON
    try:
        payload = extract_json(raw)
    except Exception as e:
        logger.error("Failed to extract JSON from LLM output: %s", e)
        fallback = heuristic_parser(subject, body, headers)
        fallback["confidence"] = 0.0
        fallback["llm_used"] = False
        return fallback

    # Payload expected to be list
    if isinstance(payload, dict):
        payload = [payload]

    validated = []

    for item in payload:
        try:
            t = TicketOut(**item)
            validated.append(t.dict())
        except ValidationError as e:
            logger.error("LLM output validation error: %s", e)
            fallback = heuristic_parser(subject, body, headers)
            fallback["confidence"] = 0.0
            fallback["llm_used"] = False
            return fallback

    # Return single ticket or list
    if len(validated) == 1:
        validated[0]["llm_used"] = True
        return validated[0]

    # MULTI-TICKET EMAIL
    # Caller must iterate and create multiple tickets
    return {
        "multiple_tickets": True,
        "tickets": validated,
        "llm_used": True
    }
