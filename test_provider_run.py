# test_provider_run.py
import os
os.environ.setdefault("LLM_PROVIDER", "cloud")

from parser import llm_provider

print("PROVIDER env:", os.getenv("LLM_PROVIDER"))
print("callable:", callable(llm_provider.parse_email_with_llm))

# quick signature check (do NOT actually call heavy LLM here)
# If you want to run a real parse, uncomment the block below and ensure provider is configured.
# subject = "VPN Access Required"
# body = "Please install OpenVPN Connect on my dev machine. Thanks, Jeeva"
# headers = "From: jeeva.s@example.com"
# res = llm_provider.parse_email_with_llm(subject, body, headers, message_id="demo-123")
# print(res)
