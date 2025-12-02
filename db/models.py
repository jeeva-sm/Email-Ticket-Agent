# db/models.py
from __future__ import annotations
import os
from datetime import datetime
from typing import Optional

from sqlalchemy import (
    Column, Integer, String, Float, Boolean, DateTime, create_engine, Text
)
from sqlalchemy.orm import declarative_base, sessionmaker

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///tickets.db")

engine = create_engine(DATABASE_URL, echo=False, future=True)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
Base = declarative_base()

class Ticket(Base):
    __tablename__ = "tickets"

    id = Column(Integer, primary_key=True, index=True)
    subject = Column(String(512), nullable=False)
    description = Column(Text, nullable=False)
    reporter_email = Column(String(256), nullable=True)
    priority = Column(String(32), nullable=False, default="medium")
    device = Column(String(256), nullable=True)
    location = Column(String(256), nullable=True)
    tags = Column(String(1024), nullable=True)  # comma-separated
    suggested_actions = Column(String(2048), nullable=True)
    confidence = Column(Float, nullable=False, default=0.0)
    llm_used = Column(Boolean, nullable=False, default=False)

    # idempotency: outlook EntryID or IMAP UID
    message_id = Column(String(512), nullable=True, unique=True, index=True)

    # lifecycle fields
    status = Column(String(32), nullable=False, default="active")  # 'active' or 'closed'
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    closed_at = Column(DateTime, nullable=True)
    closed_by = Column(String(256), nullable=True)  # user who closed via dashboard

    def __repr__(self) -> str:
        return f"<Ticket id={self.id} subject={self.subject!r} status={self.status}>"

def get_session():
    return SessionLocal()

def init_db():
    Base.metadata.create_all(bind=engine)
