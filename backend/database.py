from sqlalchemy import create_engine, Column, Integer, String, Text, DateTime, Float, JSON
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from datetime import datetime
import json

# Create SQLite database
DATABASE_URL = "sqlite:///./support_tickets.db"
engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()

class SupportTicket(Base):
    __tablename__ = "support_tickets"
    
    # Primary key - Ticket number
    ticket_number = Column(String, primary_key=True, index=True)  # TKT-000001
    
    # Email information
    sender_email = Column(String, nullable=False, index=True)
    sender_name = Column(String, nullable=True)
    subject = Column(String, nullable=True)
    body = Column(Text, nullable=False)
    email_hash = Column(String, unique=True, index=True)  # For deduplication
    
    # Ticket status
    status = Column(String, default="open", index=True)  # open, in-progress, closed
    created_at = Column(DateTime, default=datetime.utcnow, index=True)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # AI Analysis results
    summary = Column(Text)
    key_points = Column(Text)  # JSON array stored as text
    category = Column(String, index=True)
    priority = Column(String, index=True)
    sentiment_tone = Column(String)
    sentiment_confidence = Column(Float)
    suggested_reply = Column(Text)
    
    # Metadata
    word_count = Column(Integer)
    email_snippet = Column(Text, nullable=True)
    confirmation_sent = Column(DateTime, nullable=True)  # When confirmation email was sent
    
    def to_dict(self):
        """Convert to dictionary for API responses"""
        return {
            "ticket_number": self.ticket_number,
            "sender_email": self.sender_email,
            "sender_name": self.sender_name,
            "subject": self.subject,
            "status": self.status,
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "updated_at": self.updated_at.isoformat() if self.updated_at else None,
            "summary": self.summary,
            "key_points": json.loads(self.key_points) if self.key_points else [],
            "category": self.category,
            "priority": self.priority,
            "sentiment": {
                "tone": self.sentiment_tone,
                "confidence": self.sentiment_confidence
            },
            "suggested_reply": self.suggested_reply,
            "metadata": {
                "word_count": self.word_count,
                "email_snippet": self.email_snippet,
                "confirmation_sent": self.confirmation_sent.isoformat() if self.confirmation_sent else None
            }
        }
    
    def to_dict_full(self):
        """Convert to dictionary with full email body"""
        data = self.to_dict()
        data["body"] = self.body
        return data

class EmailAnalysis(Base):
    __tablename__ = "email_analyses"
    
    id = Column(Integer, primary_key=True, index=True)
    email_hash = Column(String, unique=True, index=True)  # Hash of subject+sender+date
    sender = Column(String, nullable=True)
    subject = Column(String, nullable=True)
    analyzed_at = Column(DateTime, default=datetime.utcnow)
    
    # Analysis results
    summary = Column(Text)
    key_points = Column(Text)  # JSON array stored as text
    category = Column(String)
    priority = Column(String)
    sentiment_tone = Column(String)
    sentiment_confidence = Column(Float)
    reply = Column(Text)
    
    # Metadata
    word_count = Column(Integer)
    email_snippet = Column(Text, nullable=True)  # First 200 chars of email
    
    def to_dict(self):
        """Convert to dictionary for API responses"""
        return {
            "id": self.id,
            "email_hash": self.email_hash,
            "sender": self.sender,
            "subject": self.subject,
            "analyzed_at": self.analyzed_at.isoformat() if self.analyzed_at else None,
            "summary": self.summary,
            "key_points": json.loads(self.key_points) if self.key_points else [],
            "category": self.category,
            "priority": self.priority,
            "sentiment": {
                "tone": self.sentiment_tone,
                "confidence": self.sentiment_confidence
            },
            "reply": self.reply,
            "metadata": {
                "word_count": self.word_count,
                "email_snippet": self.email_snippet
            }
        }

# Create tables
Base.metadata.create_all(bind=engine)

def get_db():
    """Dependency for database sessions"""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
