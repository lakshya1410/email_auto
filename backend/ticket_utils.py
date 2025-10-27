"""
Utility functions for ticket management system
"""
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from sqlalchemy.orm import Session
from database import SupportTicket
import os
from dotenv import load_dotenv

load_dotenv()

def generate_ticket_number(db: Session) -> str:
    """
    Generate next ticket number in format TKT-000001
    """
    # Get the latest ticket number
    latest_ticket = db.query(SupportTicket).order_by(
        SupportTicket.ticket_number.desc()
    ).first()
    
    if not latest_ticket:
        # First ticket
        return "TKT-000001"
    
    # Extract number from latest ticket (TKT-000001 -> 1)
    try:
        latest_num = int(latest_ticket.ticket_number.split('-')[1])
        next_num = latest_num + 1
        return f"TKT-{next_num:06d}"
    except (IndexError, ValueError):
        # Fallback if format is unexpected
        return "TKT-000001"


def send_confirmation_email(ticket: SupportTicket) -> bool:
    """
    Send ticket confirmation email to customer using Microsoft SMTP
    
    Returns True if successful, False otherwise
    """
    try:
        # Get email credentials from environment
        smtp_email = os.getenv("SMTP_EMAIL")
        smtp_password = os.getenv("SMTP_PASSWORD")
        
        if not smtp_email or not smtp_password:
            print("⚠️ WARNING: SMTP credentials not found in .env file")
            print("   Add SMTP_EMAIL and SMTP_PASSWORD to enable email sending")
            return False
        
        # Create email message
        msg = MIMEMultipart('alternative')
        msg['From'] = smtp_email
        msg['To'] = ticket.sender_email
        msg['Subject'] = f"Your Support Ticket #{ticket.ticket_number} - Confirmed"
        
        # Create email body
        priority_emoji = {
            "High": "🔴",
            "Medium": "🟡", 
            "Low": "🟢"
        }.get(ticket.priority, "⚪")
        
        # Parse sender name from email or use email
        sender_name = ticket.sender_name or ticket.sender_email.split('@')[0].title()
        
        # Format key points as bullet list
        key_points_list = ""
        if ticket.key_points:
            import json
            points = json.loads(ticket.key_points) if isinstance(ticket.key_points, str) else ticket.key_points
            key_points_list = "\n".join([f"• {point}" for point in points[:3]])
        
        # Create plain text version
        text_body = f"""Hello {sender_name},

Thank you for contacting us! Your support request has been received and a ticket has been created.

TICKET DETAILS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Ticket Number:  {ticket.ticket_number}
Status:         {ticket.status.title()}
Priority:       {priority_emoji} {ticket.priority}
Category:       {ticket.category}

Summary:
{ticket.summary}

{key_points_list if key_points_list else ''}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Our team is reviewing your request and will get back to you shortly.

Please reference ticket #{ticket.ticket_number} in any follow-up communications.

Best regards,
Support Team
"""
        
        # Create HTML version
        html_body = f"""
<!DOCTYPE html>
<html>
<head>
    <style>
        body {{
            font-family: 'Segoe UI', Arial, sans-serif;
            line-height: 1.6;
            color: #323130;
            max-width: 600px;
            margin: 0 auto;
            padding: 20px;
        }}
        .header {{
            background: linear-gradient(135deg, #0078d4 0%, #106ebe 100%);
            color: white;
            padding: 20px;
            border-radius: 8px 8px 0 0;
            text-align: center;
        }}
        .content {{
            background: #f9f9f9;
            padding: 30px;
            border: 1px solid #e1dfdd;
            border-top: none;
        }}
        .ticket-box {{
            background: white;
            border: 2px solid #0078d4;
            border-radius: 8px;
            padding: 20px;
            margin: 20px 0;
        }}
        .ticket-row {{
            display: flex;
            justify-content: space-between;
            padding: 8px 0;
            border-bottom: 1px solid #e1dfdd;
        }}
        .ticket-row:last-child {{
            border-bottom: none;
        }}
        .label {{
            font-weight: 600;
            color: #605e5c;
        }}
        .value {{
            color: #323130;
        }}
        .summary-box {{
            background: #fff4ce;
            border-left: 4px solid #faa300;
            padding: 15px;
            margin: 15px 0;
            border-radius: 4px;
        }}
        .key-points {{
            background: white;
            padding: 15px;
            border-radius: 4px;
            margin: 15px 0;
        }}
        .key-points ul {{
            margin: 10px 0;
            padding-left: 20px;
        }}
        .key-points li {{
            margin: 8px 0;
        }}
        .footer {{
            background: #f3f2f1;
            padding: 20px;
            border-radius: 0 0 8px 8px;
            text-align: center;
            color: #605e5c;
            font-size: 13px;
        }}
        .priority-high {{ color: #d13438; }}
        .priority-medium {{ color: #faa300; }}
        .priority-low {{ color: #107c10; }}
    </style>
</head>
<body>
    <div class="header">
        <h1 style="margin: 0;">🎫 Support Ticket Confirmation</h1>
        <p style="margin: 10px 0 0 0; opacity: 0.9;">Your request has been received</p>
    </div>
    
    <div class="content">
        <p>Hello <strong>{sender_name}</strong>,</p>
        
        <p>Thank you for contacting us! Your support request has been received and a ticket has been created.</p>
        
        <div class="ticket-box">
            <h2 style="margin-top: 0; color: #0078d4;">📋 Ticket Details</h2>
            
            <div class="ticket-row">
                <span class="label">Ticket Number:</span>
                <span class="value"><strong>{ticket.ticket_number}</strong></span>
            </div>
            
            <div class="ticket-row">
                <span class="label">Status:</span>
                <span class="value">{ticket.status.title()}</span>
            </div>
            
            <div class="ticket-row">
                <span class="label">Priority:</span>
                <span class="value priority-{ticket.priority.lower()}">{priority_emoji} {ticket.priority}</span>
            </div>
            
            <div class="ticket-row">
                <span class="label">Category:</span>
                <span class="value">{ticket.category}</span>
            </div>
        </div>
        
        <div class="summary-box">
            <strong>📝 Summary:</strong><br>
            {ticket.summary}
        </div>
        
        {f'''
        <div class="key-points">
            <strong>🔑 Key Points:</strong>
            <ul>
                {"".join([f"<li>{point}</li>" for point in (json.loads(ticket.key_points) if isinstance(ticket.key_points, str) else ticket.key_points)[:3]])}
            </ul>
        </div>
        ''' if key_points_list else ''}
        
        <p>Our team is reviewing your request and will get back to you shortly based on the priority level.</p>
        
        <p><strong>Please reference ticket #{ticket.ticket_number} in any follow-up communications.</strong></p>
    </div>
    
    <div class="footer">
        <p>This is an automated message. Please do not reply to this email.</p>
        <p>For urgent matters, please call our support line.</p>
    </div>
</body>
</html>
"""
        
        # Attach both versions
        part1 = MIMEText(text_body, 'plain')
        part2 = MIMEText(html_body, 'html')
        msg.attach(part1)
        msg.attach(part2)
        
        # Connect to Microsoft SMTP server
        print(f"📧 Sending confirmation email to {ticket.sender_email}...")
        
        # Microsoft/Outlook SMTP settings
        smtp_server = "smtp-mail.outlook.com"
        smtp_port = 587
        
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_email, smtp_password)
            server.send_message(msg)
        
        print(f"✅ Confirmation email sent successfully to {ticket.sender_email}")
        return True
        
    except Exception as e:
        print(f"❌ Failed to send confirmation email: {str(e)}")
        return False


def format_ticket_for_display(ticket: SupportTicket) -> str:
    """
    Format ticket information for console display
    """
    status_emoji = {
        "open": "🟢",
        "in-progress": "🟡",
        "closed": "⚪"
    }.get(ticket.status, "🔵")
    
    priority_emoji = {
        "High": "🔴",
        "Medium": "🟡",
        "Low": "🟢"
    }.get(ticket.priority, "⚪")
    
    return f"""
{status_emoji} Ticket {ticket.ticket_number}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
From:     {ticket.sender_email}
Subject:  {ticket.subject}
Status:   {ticket.status.title()}
Priority: {priority_emoji} {ticket.priority}
Category: {ticket.category}
Created:  {ticket.created_at.strftime('%Y-%m-%d %H:%M:%S')}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
