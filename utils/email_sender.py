"""
Email functionality for sending project estimates.
"""
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

try:
    import ssl
    EMAIL_AVAILABLE = True
except ImportError:
    EMAIL_AVAILABLE = False

def get_email_settings():
    """Get email settings from config file or return defaults."""
    config_path = os.path.join(".files", "email_config.json")
    
    if os.path.exists(config_path):
        try:
            import json
            with open(config_path, 'r') as f:
                return json.load(f)
        except:
            pass
    
    # Return default/empty settings
    return {
        "smtp_server": "",
        "smtp_port": 587,
        "sender_email": "",
        "sender_password": "",
        "sender_name": ""
    }

def save_email_settings(settings):
    """Save email settings to config file."""
    os.makedirs(".files", exist_ok=True)
    config_path = os.path.join(".files", "email_config.json")
    
    try:
        import json
        with open(config_path, 'w') as f:
            json.dump(settings, f, indent=2)
        return True
    except Exception as e:
        print(f"❌ Error saving email settings: {e}")
        return False

def send_estimate_email(pdf_path, recipient_email, recipient_name="", subject="", body=""):
    """
    Send an estimate PDF via email.
    
    Args:
        pdf_path: Path to the PDF file to send
        recipient_email: Email address of recipient
        recipient_name: Name of recipient (optional)
        subject: Email subject (optional, will use default if not provided)
        body: Email body text (optional, will use default if not provided)
    
    Returns:
        True if successful, False otherwise
    """
    if not EMAIL_AVAILABLE:
        raise ImportError("Email functionality requires ssl module")
    
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")
    
    # Get email settings
    settings = get_email_settings()
    
    if not settings.get("smtp_server") or not settings.get("sender_email"):
        raise ValueError("Email settings not configured. Please configure SMTP settings first.")
    
    # Create message
    msg = MIMEMultipart()
    msg['From'] = f"{settings.get('sender_name', 'Estimation System')} <{settings['sender_email']}>"
    msg['To'] = recipient_email
    msg['Subject'] = subject or f"Project Estimate - {datetime.now().strftime('%B %d, %Y')}"
    
    # Email body
    if not body:
        body = f"""
Dear {recipient_name or 'Valued Customer'},

Please find attached the project estimate for your review.

This estimate is valid for 30 days from the date of issue.

If you have any questions or would like to discuss this estimate, please don't hesitate to contact us.

Best regards,
{settings.get('sender_name', 'Estimation Team')}
"""
    
    msg.attach(MIMEText(body, 'plain'))
    
    # Attach PDF
    try:
        with open(pdf_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        
        encoders.encode_base64(part)
        filename = os.path.basename(pdf_path)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename= {filename}',
        )
        msg.attach(part)
    except Exception as e:
        raise Exception(f"Error attaching PDF: {e}")
    
    # Send email
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP(settings['smtp_server'], settings.get('smtp_port', 587)) as server:
            server.starttls(context=context)
            server.login(settings['sender_email'], settings['sender_password'])
            server.send_message(msg)
        
        print(f"✅ Email sent successfully to {recipient_email}")
        return True
    except smtplib.SMTPAuthenticationError:
        raise Exception("Email authentication failed. Please check your email settings.")
    except smtplib.SMTPException as e:
        raise Exception(f"Error sending email: {e}")
    except Exception as e:
        raise Exception(f"Unexpected error sending email: {e}")

