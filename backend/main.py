from fastapi import FastAPI, HTTPException, Depends, Request, Response
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.orm import Session
import os
import json
import hashlib
import google.generativeai as genai
from dotenv import load_dotenv
from database import get_db, EmailAnalysis, SupportTicket
from ticket_utils import generate_ticket_number, send_confirmation_email, format_ticket_for_display
from typing import List, Optional, Dict, Any
from datetime import datetime, timedelta
import asyncio
from concurrent.futures import ThreadPoolExecutor

# Load environment variables
load_dotenv()

app = FastAPI()

# Configure Gemini API
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if not GEMINI_API_KEY:
    print("⚠️  WARNING: GEMINI_API_KEY not found in environment variables!")
    print("   Create a .env file with your API key to enable AI summarization")
else:
    genai.configure(api_key=GEMINI_API_KEY)
    print("✅ Gemini API configured successfully")

# Allow all origins for development (Outlook, browsers, etc.)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allow all origins
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class EmailText(BaseModel):
    text: str
    sender: Optional[str] = None
    subject: Optional[str] = None

def generate_email_hash(text: str, sender: str = None, subject: str = None) -> str:
    """Generate unique hash for email to prevent duplicates"""
    content = f"{text[:500]}{sender or ''}{subject or ''}"
    return hashlib.md5(content.encode()).hexdigest()

@app.post("/api/summarize")
def summarize(email: EmailText, db: Session = Depends(get_db)):
    text = email.text or ""
    sender = email.sender
    subject = email.subject
    
    if not text.strip():
        return {
            "summary": "No email body detected.",
            "key_points": ["No content to analyze"],
            "category": "General",
            "priority": "Low",
            "sentiment": {"tone": "Neutral", "confidence": 0.0},
            "reply": "Unable to generate reply - no email content found.",
            "metadata": {
                "analyzed_at": __import__('datetime').datetime.now().isoformat(),
                "word_count": 0
            }
        }
    
    # Check if API key is configured
    if not GEMINI_API_KEY:
        # Fallback to basic extraction if no API key
        import re
        snippet = text.strip().replace("\r", " ").replace("\n", " ")
        sentences = re.split(r'(?<=[.!?])\s+', snippet)
        summary = " ".join(sentences[:3])
        if len(summary) > 800:
            summary = summary[:800] + "..."
        return {
            "summary": f"⚠️ [Fallback Mode - No API Key]\n\n{summary}\n\nTo enable AI summarization, add your GEMINI_API_KEY to the .env file.",
            "key_points": ["API key required for analysis"],
            "category": "General",
            "priority": "Medium",
            "sentiment": {"tone": "Neutral", "confidence": 0.0},
            "reply": "⚠️ API key required to generate suggested replies.",
            "metadata": {
                "analyzed_at": __import__('datetime').datetime.now().isoformat(),
                "word_count": len(text.split())
            }
        }
    
    try:
        # Initialize Gemini model
        model = genai.GenerativeModel('gemini-2.5-flash')
        
        # OPTIMIZED: Single comprehensive prompt for all analysis
        comprehensive_prompt = f"""Analyze the following email and provide a comprehensive analysis in JSON format.

Email:
{text}

Please analyze and return ONLY valid JSON with this exact structure:
{{
    "summary": "3-5 bullet points summarizing the email",
    "key_points": ["specific date/name/number", "action item", "critical info"],
    "category": "Sales|Support|General|Marketing|HR",
    "priority": "High|Medium|Low",
    "sentiment": {{"tone": "Positive|Neutral|Negative|Urgent", "confidence": 0.95}},
    "reply": "Professional 2-4 paragraph reply with greeting and closing"
}}

Categories guide:
- Sales: proposals, pricing, purchases, orders, deals
- Support: issues, problems, help, bugs, technical questions
- General: general communication, updates, casual messages
- Marketing: campaigns, newsletters, promotions, webinars
- HR: recruitment, interviews, performance, team matters

Priority guide:
- High: urgent, deadline-driven, requires immediate action
- Medium: important but not urgent, can wait 1-2 days
- Low: informational, FYI, no immediate action needed

Sentiment guide:
- Positive: friendly, thankful, enthusiastic
- Neutral: informational, professional
- Negative: complaint, frustration, anger
- Urgent: time-sensitive, requires immediate attention"""
        
        print("🤖 Generating comprehensive analysis...")
        response = model.generate_content(comprehensive_prompt)
        
        if not response or not response.text:
            raise Exception("Empty response from Gemini API")
        
        response_text = response.text.strip()
        
        # Clean up response - extract JSON if wrapped in markdown
        if "```json" in response_text:
            response_text = response_text.split("```json")[1].split("```")[0].strip()
        elif "```" in response_text:
            response_text = response_text.split("```")[1].split("```")[0].strip()
        
        # Parse JSON response
        analysis = json.loads(response_text)
        
        # Validate and set defaults
        if "summary" not in analysis:
            analysis["summary"] = "Unable to generate summary"
        if "key_points" not in analysis or not isinstance(analysis["key_points"], list):
            analysis["key_points"] = ["Unable to extract key points"]
        if "category" not in analysis or analysis["category"] not in ["Sales", "Support", "General", "Marketing", "HR"]:
            analysis["category"] = "General"
        if "priority" not in analysis or analysis["priority"] not in ["High", "Medium", "Low"]:
            analysis["priority"] = "Medium"
        if "sentiment" not in analysis or "tone" not in analysis["sentiment"]:
            analysis["sentiment"] = {"tone": "Neutral", "confidence": 0.5}
        if "reply" not in analysis:
            analysis["reply"] = "Unable to generate reply"
        
        # Add metadata
        analysis["metadata"] = {
            "analyzed_at": __import__('datetime').datetime.now().isoformat(),
            "word_count": len(text.split())
        }
        
        print("✅ Analysis complete!")
        
        # Save to database
        try:
            email_hash = generate_email_hash(text, sender, subject)
            
            # Check if already analyzed
            existing = db.query(EmailAnalysis).filter(EmailAnalysis.email_hash == email_hash).first()
            
            if existing:
                print(f"📌 Updating existing analysis (ID: {existing.id})")
                existing.summary = analysis["summary"]
                existing.key_points = json.dumps(analysis["key_points"])
                existing.category = analysis["category"]
                existing.priority = analysis["priority"]
                existing.sentiment_tone = analysis["sentiment"]["tone"]
                existing.sentiment_confidence = analysis["sentiment"]["confidence"]
                existing.reply = analysis["reply"]
                existing.word_count = len(text.split())
                existing.analyzed_at = datetime.now()
                db.commit()
                analysis["metadata"]["database_id"] = existing.id
            else:
                print("💾 Saving new analysis to database")
                db_analysis = EmailAnalysis(
                    email_hash=email_hash,
                    sender=sender,
                    subject=subject,
                    summary=analysis["summary"],
                    key_points=json.dumps(analysis["key_points"]),
                    category=analysis["category"],
                    priority=analysis["priority"],
                    sentiment_tone=analysis["sentiment"]["tone"],
                    sentiment_confidence=analysis["sentiment"]["confidence"],
                    reply=analysis["reply"],
                    word_count=len(text.split()),
                    email_snippet=text[:200]
                )
                db.add(db_analysis)
                db.commit()
                db.refresh(db_analysis)
                analysis["metadata"]["database_id"] = db_analysis.id
                print(f"✅ Saved to database (ID: {db_analysis.id})")
        except Exception as db_error:
            print(f"⚠️  Database error: {str(db_error)}")
            # Continue even if database save fails
        
        return analysis
    
    except json.JSONDecodeError as e:
        print(f"❌ JSON parsing error: {str(e)}")
        print(f"Response was: {response_text[:200]}...")
        
        # Try to extract what we can from the response
        return {
            "summary": response_text[:500] if response_text else "Error parsing response",
            "key_points": ["Unable to parse structured data"],
            "category": "General",
            "priority": "Medium",
            "sentiment": {"tone": "Neutral", "confidence": 0.0},
            "reply": "Unable to generate structured reply",
            "metadata": {
                "analyzed_at": __import__('datetime').datetime.now().isoformat(),
                "word_count": len(text.split()),
                "error": "JSON parse error"
            }
        }
    
    except Exception as e:
        print(f"❌ Error calling Gemini API: {str(e)}")
        
        # Fallback to basic extraction on error
        import re
        snippet = text.strip().replace("\r", " ").replace("\n", " ")
        sentences = re.split(r'(?<=[.!?])\s+', snippet)
        fallback_summary = " ".join(sentences[:3])
        if len(fallback_summary) > 800:
            fallback_summary = fallback_summary[:800] + "..."
        
        return {
            "summary": f"⚠️ [Error: {str(e)}]\n\nFallback summary:\n{fallback_summary}",
            "key_points": ["Analysis unavailable due to error"],
            "category": "General",
            "priority": "Medium",
            "sentiment": {"tone": "Neutral", "confidence": 0.0},
            "reply": f"⚠️ Unable to generate reply due to error: {str(e)}",
            "metadata": {
                "analyzed_at": __import__('datetime').datetime.now().isoformat(),
                "word_count": len(text.split()),
                "error": str(e)
            }
        }

@app.get("/")
def root():
    return {
        "message": "Email Summarizer API - Enhanced Edition",
        "status": "running",
        "gemini_configured": bool(GEMINI_API_KEY),
        "features": [
            "Smart summarization",
            "Key points extraction",
            "Category detection",
            "Priority classification",
            "Sentiment analysis",
            "Reply generation",
            "Analysis history",
            "Dashboard analytics"
        ],
        "endpoints": {
            "analyze": "/api/summarize (POST)",
            "history": "/api/history (GET)",
            "stats": "/api/stats (GET)",
            "delete": "/api/history/{id} (DELETE)",
            "docs": "/docs"
        }
    }

@app.get("/api/history")
def get_history(
    limit: int = 50,
    offset: int = 0,
    category: Optional[str] = None,
    priority: Optional[str] = None,
    db: Session = Depends(get_db)
):
    """Get analysis history with optional filters"""
    query = db.query(EmailAnalysis)
    
    # Apply filters
    if category:
        query = query.filter(EmailAnalysis.category == category)
    if priority:
        query = query.filter(EmailAnalysis.priority == priority)
    
    # Order by most recent first
    query = query.order_by(EmailAnalysis.analyzed_at.desc())
    
    # Pagination
    total = query.count()
    analyses = query.offset(offset).limit(limit).all()
    
    return {
        "total": total,
        "limit": limit,
        "offset": offset,
        "items": [analysis.to_dict() for analysis in analyses]
    }

@app.get("/api/history/{analysis_id}")
def get_analysis(analysis_id: int, db: Session = Depends(get_db)):
    """Get specific analysis by ID"""
    analysis = db.query(EmailAnalysis).filter(EmailAnalysis.id == analysis_id).first()
    
    if not analysis:
        raise HTTPException(status_code=404, detail="Analysis not found")
    
    return analysis.to_dict()

@app.delete("/api/history/{analysis_id}")
def delete_analysis(analysis_id: int, db: Session = Depends(get_db)):
    """Delete specific analysis"""
    analysis = db.query(EmailAnalysis).filter(EmailAnalysis.id == analysis_id).first()
    
    if not analysis:
        raise HTTPException(status_code=404, detail="Analysis not found")
    
    db.delete(analysis)
    db.commit()
    
    return {"message": "Analysis deleted successfully", "id": analysis_id}

@app.get("/api/stats")
def get_stats(days: int = 30, db: Session = Depends(get_db)):
    """Get dashboard statistics"""
    from sqlalchemy import func
    
    # Calculate date range
    start_date = datetime.now() - timedelta(days=days)
    
    # Total emails analyzed
    total_analyzed = db.query(EmailAnalysis).filter(
        EmailAnalysis.analyzed_at >= start_date
    ).count()
    
    # By category
    by_category = db.query(
        EmailAnalysis.category,
        func.count(EmailAnalysis.id).label('count')
    ).filter(
        EmailAnalysis.analyzed_at >= start_date
    ).group_by(EmailAnalysis.category).all()
    
    # By priority
    by_priority = db.query(
        EmailAnalysis.priority,
        func.count(EmailAnalysis.id).label('count')
    ).filter(
        EmailAnalysis.analyzed_at >= start_date
    ).group_by(EmailAnalysis.priority).all()
    
    # By sentiment
    by_sentiment = db.query(
        EmailAnalysis.sentiment_tone,
        func.count(EmailAnalysis.id).label('count')
    ).filter(
        EmailAnalysis.analyzed_at >= start_date
    ).group_by(EmailAnalysis.sentiment_tone).all()
    
    # Average sentiment confidence
    avg_sentiment = db.query(
        func.avg(EmailAnalysis.sentiment_confidence)
    ).filter(
        EmailAnalysis.analyzed_at >= start_date
    ).scalar() or 0
    
    # Recent analyses (last 7 days by day)
    recent = db.query(
        func.date(EmailAnalysis.analyzed_at).label('date'),
        func.count(EmailAnalysis.id).label('count')
    ).filter(
        EmailAnalysis.analyzed_at >= datetime.now() - timedelta(days=7)
    ).group_by(func.date(EmailAnalysis.analyzed_at)).all()
    
    return {
        "period_days": days,
        "total_analyzed": total_analyzed,
        "by_category": {cat: count for cat, count in by_category},
        "by_priority": {pri: count for pri, count in by_priority},
        "by_sentiment": {sent: count for sent, count in by_sentiment},
        "avg_sentiment_confidence": round(float(avg_sentiment), 2),
        "recent_activity": [
            {"date": str(date), "count": count} for date, count in recent
        ]
    }


# ============================================
# TICKET MANAGEMENT SYSTEM ENDPOINTS
# ============================================

class TicketCreate(BaseModel):
    sender_email: str
    sender_name: Optional[str] = None
    subject: str
    body: str

class TicketStatusUpdate(BaseModel):
    status: str  # open, in-progress, closed

@app.post("/api/tickets/create")
def create_ticket(ticket_data: TicketCreate, db: Session = Depends(get_db)):
    """
    Create a new support ticket from email
    Automatically analyzes content and sends confirmation email
    """
    try:
        # Check for duplicate based on email hash
        email_hash = generate_email_hash(
            ticket_data.body,
            ticket_data.sender_email,
            ticket_data.subject
        )
        
        existing_ticket = db.query(SupportTicket).filter(
            SupportTicket.email_hash == email_hash
        ).first()
        
        if existing_ticket:
            print(f"⚠️ Duplicate ticket detected: {existing_ticket.ticket_number}")
            return {
                "ticket_number": existing_ticket.ticket_number,
                "status": "duplicate",
                "message": "A ticket for this email already exists",
                "existing_ticket": existing_ticket.to_dict()
            }
        
        # Generate ticket number
        ticket_number = generate_ticket_number(db)
        print(f"\n🎫 Creating ticket {ticket_number}...")
        
        # Analyze email content using AI
        print("🤖 Analyzing email content...")
        text = ticket_data.body
        
        if not GEMINI_API_KEY:
            raise HTTPException(status_code=500, detail="AI API key not configured")
        
        model = genai.GenerativeModel("gemini-2.0-flash-exp")
        
        comprehensive_prompt = f"""Analyze this customer support email and provide structured information.

Email:
{text}

Please analyze and return ONLY valid JSON with this exact structure:
{{
    "summary": "3-5 bullet points summarizing the email",
    "key_points": ["specific date/name/number", "action item", "critical info"],
    "category": "Sales|Support|General|Marketing|HR",
    "priority": "High|Medium|Low",
    "sentiment": {{"tone": "Positive|Neutral|Negative|Urgent", "confidence": 0.95}},
    "reply": "Professional 2-4 paragraph reply with greeting and closing"
}}

Categories guide:
- Sales: proposals, pricing, purchases, orders, deals
- Support: issues, problems, help, bugs, technical questions
- General: general communication, updates, casual messages
- Marketing: campaigns, newsletters, promotions, webinars
- HR: recruitment, interviews, performance, team matters

Priority guide:
- High: urgent, deadline-driven, requires immediate action
- Medium: important but not urgent, can wait 1-2 days
- Low: informational, FYI, no immediate action needed

Sentiment guide:
- Positive: friendly, thankful, enthusiastic
- Neutral: informational, professional
- Negative: complaint, frustration, anger
- Urgent: time-sensitive, requires immediate attention"""
        
        response = model.generate_content(comprehensive_prompt)
        
        if not response or not response.text:
            raise HTTPException(status_code=500, detail="AI analysis failed")
        
        response_text = response.text.strip()
        
        # Clean up response
        if "```json" in response_text:
            response_text = response_text.split("```json")[1].split("```")[0].strip()
        elif "```" in response_text:
            response_text = response_text.split("```")[1].split("```")[0].strip()
        
        analysis = json.loads(response_text)
        
        # Create ticket in database
        word_count = len(text.split())
        email_snippet = text[:200] + "..." if len(text) > 200 else text
        
        # Extract and format analysis fields properly
        summary = analysis.get("summary", "No summary")
        if isinstance(summary, list):
            summary = "\n".join(f"• {point}" for point in summary)
        
        key_points = analysis.get("key_points", [])
        if not isinstance(key_points, list):
            key_points = [str(key_points)]
        
        new_ticket = SupportTicket(
            ticket_number=ticket_number,
            sender_email=ticket_data.sender_email,
            sender_name=ticket_data.sender_name,
            subject=ticket_data.subject,
            body=ticket_data.body,
            email_hash=email_hash,
            status="open",
            summary=summary,
            key_points=json.dumps(key_points),
            category=analysis.get("category", "General"),
            priority=analysis.get("priority", "Medium"),
            sentiment_tone=analysis.get("sentiment", {}).get("tone", "Neutral"),
            sentiment_confidence=analysis.get("sentiment", {}).get("confidence", 0.0),
            suggested_reply=analysis.get("reply", ""),
            word_count=word_count,
            email_snippet=email_snippet
        )
        
        db.add(new_ticket)
        db.commit()
        db.refresh(new_ticket)
        
        print(format_ticket_for_display(new_ticket))
        print("✅ Ticket created successfully!")
        
        # Send confirmation email
        print("📧 Sending confirmation email...")
        email_sent = send_confirmation_email(new_ticket)
        
        if email_sent:
            new_ticket.confirmation_sent = datetime.utcnow()
            db.commit()
        
        return {
            "ticket_number": ticket_number,
            "status": "created",
            "message": "Ticket created successfully",
            "confirmation_sent": email_sent,
            "ticket": new_ticket.to_dict()
        }
        
    except json.JSONDecodeError as e:
        print(f"❌ JSON parsing error: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Failed to parse AI response: {str(e)}")
    except Exception as e:
        print(f"❌ Error creating ticket: {str(e)}")
        db.rollback()
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/tickets")
def get_tickets(
    limit: int = 50,
    offset: int = 0,
    status: Optional[str] = None,
    category: Optional[str] = None,
    priority: Optional[str] = None,
    db: Session = Depends(get_db)
):
    """Get list of all tickets with optional filters, sorted by most recent"""
    query = db.query(SupportTicket)
    
    # Apply filters
    if status:
        query = query.filter(SupportTicket.status == status)
    if category:
        query = query.filter(SupportTicket.category == category)
    if priority:
        query = query.filter(SupportTicket.priority == priority)
    
    # Order by most recent first
    query = query.order_by(SupportTicket.created_at.desc())
    
    # Pagination
    total = query.count()
    tickets = query.offset(offset).limit(limit).all()
    
    return {
        "total": total,
        "limit": limit,
        "offset": offset,
        "items": [ticket.to_dict() for ticket in tickets]
    }


@app.get("/api/tickets/{ticket_number}")
def get_ticket(ticket_number: str, db: Session = Depends(get_db)):
    """Get specific ticket by ticket number"""
    ticket = db.query(SupportTicket).filter(
        SupportTicket.ticket_number == ticket_number
    ).first()
    
    if not ticket:
        raise HTTPException(status_code=404, detail="Ticket not found")
    
    return ticket.to_dict_full()


@app.patch("/api/tickets/{ticket_number}/status")
def update_ticket_status(
    ticket_number: str,
    status_update: TicketStatusUpdate,
    db: Session = Depends(get_db)
):
    """Update ticket status (open, in-progress, closed)"""
    ticket = db.query(SupportTicket).filter(
        SupportTicket.ticket_number == ticket_number
    ).first()
    
    if not ticket:
        raise HTTPException(status_code=404, detail="Ticket not found")
    
    # Validate status
    valid_statuses = ["open", "in-progress", "closed"]
    if status_update.status not in valid_statuses:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid status. Must be one of: {', '.join(valid_statuses)}"
        )
    
    ticket.status = status_update.status
    ticket.updated_at = datetime.utcnow()
    db.commit()
    
    print(f"✅ Ticket {ticket_number} status updated to: {status_update.status}")
    
    return {
        "ticket_number": ticket_number,
        "status": status_update.status,
        "updated_at": ticket.updated_at.isoformat()
    }


@app.delete("/api/tickets/{ticket_number}")
def delete_ticket(ticket_number: str, db: Session = Depends(get_db)):
    """Delete a specific ticket"""
    ticket = db.query(SupportTicket).filter(
        SupportTicket.ticket_number == ticket_number
    ).first()
    
    if not ticket:
        raise HTTPException(status_code=404, detail="Ticket not found")
    
    db.delete(ticket)
    db.commit()
    
    print(f"🗑️ Ticket {ticket_number} deleted")
    
    return {"message": "Ticket deleted successfully", "ticket_number": ticket_number}


@app.get("/api/tickets/stats/dashboard")
def get_ticket_stats(days: int = 30, db: Session = Depends(get_db)):
    """Get ticket statistics for dashboard"""
    from sqlalchemy import func
    
    start_date = datetime.now() - timedelta(days=days)
    
    # Total tickets
    total_tickets = db.query(SupportTicket).filter(
        SupportTicket.created_at >= start_date
    ).count()
    
    # Open tickets
    open_tickets = db.query(SupportTicket).filter(
        SupportTicket.status == "open"
    ).count()
    
    # By status
    by_status = db.query(
        SupportTicket.status,
        func.count(SupportTicket.ticket_number).label('count')
    ).filter(
        SupportTicket.created_at >= start_date
    ).group_by(SupportTicket.status).all()
    
    # By category
    by_category = db.query(
        SupportTicket.category,
        func.count(SupportTicket.ticket_number).label('count')
    ).filter(
        SupportTicket.created_at >= start_date
    ).group_by(SupportTicket.category).all()
    
    # By priority
    by_priority = db.query(
        SupportTicket.priority,
        func.count(SupportTicket.ticket_number).label('count')
    ).filter(
        SupportTicket.created_at >= start_date
    ).group_by(SupportTicket.priority).all()
    
    # Recent activity (last 7 days)
    recent = db.query(
        func.date(SupportTicket.created_at).label('date'),
        func.count(SupportTicket.ticket_number).label('count')
    ).filter(
        SupportTicket.created_at >= datetime.now() - timedelta(days=7)
    ).group_by(func.date(SupportTicket.created_at)).all()
    
    return {
        "period_days": days,
        "total_tickets": total_tickets,
        "open_tickets": open_tickets,
        "by_status": {status: count for status, count in by_status},
        "by_category": {cat: count for cat, count in by_category},
        "by_priority": {pri: count for pri, count in by_priority},
        "recent_activity": [
            {"date": str(date), "count": count} for date, count in recent
        ]
    }


# ============================================
# MICROSOFT GRAPH WEBHOOK ENDPOINTS
# ============================================

# Thread pool for async email processing
executor = ThreadPoolExecutor(max_workers=3)

def process_email_notification_sync(message_id: str, db: Session):
    """Synchronous function to process email and create ticket"""
    from webhook_manager import GraphWebhookManager
    
    try:
        print(f"\n📧 Processing email notification: {message_id}")
        
        # Initialize webhook manager
        manager = GraphWebhookManager()
        
        # Fetch email details
        print("📥 Fetching email details from Graph API...")
        email_data = manager.get_email_details(message_id)
        
        # Extract email information
        sender_email = email_data.get("from", {}).get("emailAddress", {}).get("address", "unknown@unknown.com")
        sender_name = email_data.get("from", {}).get("emailAddress", {}).get("name", sender_email)
        subject = email_data.get("subject", "No Subject")
        body = email_data.get("body", {}).get("content", "")
        
        print(f"   From: {sender_name} <{sender_email}>")
        print(f"   Subject: {subject}")
        
        # Check for duplicates
        email_hash = hashlib.md5(f"{body[:500]}{sender_email}{subject}".encode()).hexdigest()
        existing = db.query(SupportTicket).filter(
            SupportTicket.email_hash == email_hash
        ).first()
        
        if existing:
            print(f"⚠️  Duplicate email detected - Ticket {existing.ticket_number} already exists")
            return
        
        # Generate ticket number
        ticket_number = generate_ticket_number(db)
        print(f"🎫 Generated ticket number: {ticket_number}")
        
        # Analyze email with AI
        print("🤖 Analyzing email with Gemini AI...")
        model = genai.GenerativeModel('gemini-2.0-flash-exp')
        
        comprehensive_prompt = f"""Analyze the following email and provide a comprehensive analysis in JSON format.

Email:
{body}

Please analyze and return ONLY valid JSON with this exact structure:
{{
  "summary": "3-5 bullet points summarizing the email",
  "key_points": ["fact 1", "fact 2", "fact 3"],
  "category": "Sales|Support|General|Marketing|HR",
  "priority": "High|Medium|Low",
  "sentiment_tone": "Positive|Neutral|Negative|Urgent",
  "sentiment_confidence": 0.95,
  "suggested_reply": "A professional, helpful reply email"
}}"""

        response = model.generate_content(comprehensive_prompt)
        analysis = json.loads(response.text.strip().replace("```json", "").replace("```", "").strip())
        
        print(f"   Category: {analysis.get('category', 'General')}")
        print(f"   Priority: {analysis.get('priority', 'Medium')}")
        
        # Format summary as bullet points
        summary = analysis.get("summary", "No summary")
        if isinstance(summary, list):
            summary = "\n".join(f"• {point}" for point in summary)
        
        # Ensure key_points is a list
        key_points = analysis.get("key_points", [])
        if not isinstance(key_points, list):
            key_points = [str(key_points)]
        
        # Create ticket in database
        new_ticket = SupportTicket(
            ticket_number=ticket_number,
            sender_email=sender_email,
            sender_name=sender_name,
            subject=subject,
            body=body,
            email_hash=email_hash,
            status="open",
            summary=summary,
            key_points=json.dumps(key_points),
            category=analysis.get("category", "General"),
            priority=analysis.get("priority", "Medium"),
            sentiment_tone=analysis.get("sentiment_tone", "Neutral"),
            sentiment_confidence=float(analysis.get("sentiment_confidence", 0.0)),
            suggested_reply=analysis.get("suggested_reply", ""),
            word_count=len(body.split()),
            email_snippet=body[:200] if body else "",
            created_at=datetime.utcnow(),
            updated_at=datetime.utcnow()
        )
        
        db.add(new_ticket)
        db.commit()
        db.refresh(new_ticket)
        
        print(f"✅ Ticket {ticket_number} created successfully!")
        
        # Send confirmation email
        print("📧 Sending confirmation email...")
        email_sent = send_confirmation_email(new_ticket)
        
        if email_sent:
            new_ticket.confirmation_sent = datetime.utcnow()
            db.commit()
            print("✅ Confirmation email sent!")
        else:
            print("⚠️  Failed to send confirmation email")
        
        print(f"✨ Processing complete for ticket {ticket_number}\n")
        
    except Exception as e:
        print(f"❌ Error processing email notification: {str(e)}")
        import traceback
        traceback.print_exc()
        db.rollback()


@app.post("/api/webhooks/graph-notifications")
async def handle_graph_webhook(request: Request, db: Session = Depends(get_db)):
    """
    Handle Microsoft Graph API webhook notifications for new emails
    
    This endpoint:
    1. Validates webhook subscription (on first call)
    2. Receives notifications when new emails arrive
    3. Fetches email details from Graph API
    4. Creates support tickets automatically
    5. Sends confirmation emails
    """
    
    # Handle validation request from Microsoft
    validation_token = request.query_params.get("validationToken")
    if validation_token:
        print("✅ Webhook validation request received")
        return Response(content=validation_token, media_type="text/plain")
    
    try:
        # Parse notification payload
        body = await request.json()
        
        # Microsoft sends an array of notifications
        notifications = body.get("value", [])
        
        print(f"\n📬 Received {len(notifications)} notification(s) from Microsoft Graph")
        
        for notification in notifications:
            # Verify client state for security
            client_state = notification.get("clientState")
            if client_state != "SecretClientState":
                print("⚠️  Invalid client state - ignoring notification")
                continue
            
            # Extract resource data
            resource_data = notification.get("resourceData", {})
            message_id = resource_data.get("id")
            
            if not message_id:
                print("⚠️  No message ID in notification - skipping")
                continue
            
            # Process email in background thread (non-blocking)
            loop = asyncio.get_event_loop()
            loop.run_in_executor(
                executor,
                process_email_notification_sync,
                message_id,
                db
            )
        
        return {"status": "accepted", "message": f"Processing {len(notifications)} notification(s)"}
        
    except Exception as e:
        print(f"❌ Error handling webhook: {str(e)}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/webhooks/subscriptions")
def list_webhook_subscriptions():
    """List all active Microsoft Graph webhook subscriptions"""
    try:
        from webhook_manager import GraphWebhookManager
        manager = GraphWebhookManager()
        subscriptions = manager.list_subscriptions()
        
        return {
            "count": len(subscriptions),
            "subscriptions": subscriptions
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/webhooks/subscriptions/create")
def create_webhook_subscription():
    """Create a new Microsoft Graph webhook subscription"""
    try:
        from webhook_manager import GraphWebhookManager
        manager = GraphWebhookManager()
        subscription = manager.create_subscription()
        manager.save_subscription_info()
        
        return {
            "status": "created",
            "subscription": subscription
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/webhooks/subscriptions/renew")
def renew_webhook_subscription(subscription_id: Optional[str] = None):
    """Renew an existing Microsoft Graph webhook subscription"""
    try:
        from webhook_manager import GraphWebhookManager
        manager = GraphWebhookManager()
        
        # Try to load saved subscription ID if not provided
        if not subscription_id:
            manager.load_subscription_info()
        
        subscription = manager.renew_subscription(subscription_id)
        manager.save_subscription_info()
        
        return {
            "status": "renewed",
            "subscription": subscription
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.delete("/api/webhooks/subscriptions/{subscription_id}")
def delete_webhook_subscription(subscription_id: str):
    """Delete a Microsoft Graph webhook subscription"""
    try:
        from webhook_manager import GraphWebhookManager
        manager = GraphWebhookManager()
        success = manager.delete_subscription(subscription_id)
        
        if success:
            return {"status": "deleted", "subscription_id": subscription_id}
        else:
            raise HTTPException(status_code=500, detail="Failed to delete subscription")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


