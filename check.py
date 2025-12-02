# quick check.py
import sqlite3
from urllib.parse import urlparse
import os

db = os.getenv("DATABASE_URL","sqlite:///tickets.db")
if db.startswith("sqlite:///"):
    path = db[len("sqlite:///"):]
else:
    path = db
conn = sqlite3.connect(path)
c = conn.cursor()
c.execute("PRAGMA table_info(tickets)")
print([r[1] for r in c.fetchall()])
conn.close()
