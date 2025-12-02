import os, sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from db.models import get_session, Ticket
from sqlalchemy import delete

print("Clearing tickets table...")

session = get_session()
try:
    session.execute(delete(Ticket))
    session.commit()
    print("All ticket rows deleted.")
finally:
    session.close()
