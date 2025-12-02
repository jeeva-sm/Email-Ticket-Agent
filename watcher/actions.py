# watcher/actions.py
"""
Ticket lifecycle actions used by the dashboard API and watcher.

Functions:
- close_ticket(ticket_id, closed_by, move_mail) -> bool
- reopen_ticket(ticket_id, reopened_by) -> bool
- get_ticket(ticket_id) -> Ticket | None
- get_ticket_by_message_id(message_id) -> Ticket | None
"""

from datetime import datetime
import logging
from typing import Optional

from db.models import get_session, Ticket
from watcher.outlook_utils import move_message_by_entryid

logger = logging.getLogger(__name__)


def get_ticket(ticket_id: int) -> Optional[Ticket]:
    session = get_session()
    try:
        return session.query(Ticket).filter(Ticket.id == ticket_id).first()
    finally:
        session.close()


def get_ticket_by_message_id(message_id: str) -> Optional[Ticket]:
    if not message_id:
        return None
    session = get_session()
    try:
        return session.query(Ticket).filter(Ticket.message_id == message_id).first()
    finally:
        session.close()


def close_ticket(ticket_id: int, closed_by: str = "dashboard_user", move_mail: bool = True) -> bool:
    """
    Mark ticket as closed in DB and optionally move the related mail to Processed folder.
    Returns True on success, False otherwise.
    """
    session = get_session()
    try:
        t = session.query(Ticket).filter(Ticket.id == ticket_id).with_for_update().first()
        if not t:
            logger.warning("close_ticket: ticket %s not found", ticket_id)
            return False

        if t.status == "closed":
            logger.info("close_ticket: ticket %s already closed", ticket_id)
            return True

        t.status = "closed"
        t.closed_at = datetime.utcnow()
        t.closed_by = closed_by
        session.add(t)
        session.commit()
        logger.info("close_ticket: ticket %s closed by %s", ticket_id, closed_by)

        # Attempt to move mail if requested and message_id exists
        if move_mail and t.message_id:
            moved = False
            try:
                moved = move_message_by_entryid(t.message_id)
            except Exception as e:
                logger.exception("close_ticket: moving mail failed for ticket %s: %s", ticket_id, e)
            if not moved:
                logger.warning("close_ticket: could not move mail for ticket %s (message_id=%s)", ticket_id, t.message_id)
        return True
    except Exception as e:
        session.rollback()
        logger.exception("close_ticket: DB error for ticket %s: %s", ticket_id, e)
        return False
    finally:
        session.close()


def reopen_ticket(ticket_id: int, reopened_by: str = "dashboard_user") -> bool:
    """
    Reopen a previously closed ticket. Does NOT move mail back automatically.
    """
    session = get_session()
    try:
        t = session.query(Ticket).filter(Ticket.id == ticket_id).with_for_update().first()
        if not t:
            logger.warning("reopen_ticket: ticket %s not found", ticket_id)
            return False
        if t.status != "closed":
            logger.info("reopen_ticket: ticket %s is not closed (status=%s)", ticket_id, t.status)
            return True

        t.status = "active"
        t.closed_at = None
        t.closed_by = None
        # Optional: append a note into suggested_actions or tags that it was reopened
        session.add(t)
        session.commit()
        logger.info("reopen_ticket: ticket %s reopened by %s", ticket_id, reopened_by)
        return True
    except Exception as e:
        session.rollback()
        logger.exception("reopen_ticket: DB error for ticket %s: %s", ticket_id, e)
        return False
    finally:
        session.close()
