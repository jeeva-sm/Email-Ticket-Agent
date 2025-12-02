import load_env
from parser.llm_provider import parse_email_with_llm
s = "Install AnyDesk"
b = "I have the AnyDesk installer. Please connect to install. AnyDesk ID: 1791342921"
h = "From: jeeva.s@example.com"
print(parse_email_with_llm(s, b, h, message_id="demo-ollama-1"))