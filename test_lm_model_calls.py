# test_lm_model_calls.py
from parser.lmstudio_client import call_lmstudio
print("Calling LM Studio (completions preferred)...")
resp = call_lmstudio('Return only the JSON: {"ok": true}', max_tokens=60)
print("RESPONSE:")
print(resp)