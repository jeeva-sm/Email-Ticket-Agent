# load_env.py
"""
Loads environment variables from .env automatically on import.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent
ENV_PATH = ROOT / ".env"

if ENV_PATH.exists():
    load_dotenv(dotenv_path=ENV_PATH)
else:
    print("[load_env] WARNING: .env file not found ->", ENV_PATH)
