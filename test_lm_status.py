import os
os.environ.setdefault("LLM_PROVIDER", "cloud")

from parser import llm_provider

print("PROVIDER env:", os.getenv("LLM_PROVIDER"))
print("callable:", callable(llm_provider.parse_email_with_llm))
