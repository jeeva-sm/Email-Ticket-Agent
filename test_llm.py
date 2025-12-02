# test_llm.py
import requests, os
url = os.getenv("LMSTUDIO_URL", "http://127.0.0.1:1234").rstrip("/") + "/v1/chat/completions"
print("POST ->", url)
payload = {
    "model": None,
    "messages": [
        {"role": "system", "content": "You are a JSON-only extraction assistant."},
        {"role": "user", "content": 'Hello, respond with OK.'}
    ],
    "max_tokens": 50,
    "temperature": 0.0,
}
r = requests.post(url, json=payload, timeout=30)
print("STATUS:", r.status_code)
print("RESPONSE TEXT:\n", r.text)
