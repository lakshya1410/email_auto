"""
Microsoft Graph Webhook Setup Script
Helps configure and test the webhook-based email automation
"""

import os
import sys
from dotenv import load_dotenv
from webhook_manager import GraphWebhookManager

load_dotenv()

def print_header(text):
    """Print formatted header"""
    print("\n" + "="*60)
    print(f"  {text}")
    print("="*60 + "\n")

def check_environment():
    """Verify all required environment variables are set"""
    print_header("🔍 STEP 1: Checking Environment Variables")
    
    required_vars = {
        "CLIENT_ID": "Azure App Client ID",
        "CLIENT_SECRET": "Azure App Client Secret", 
        "WEBHOOK_URL": "Public HTTPS webhook URL",
        "IMAP_EMAIL": "Email address to monitor",
        "GEMINI_API_KEY": "Google Gemini API key"
    }
    
    missing = []
    
    for var, description in required_vars.items():
        value = os.getenv(var)
        if value:
            # Mask sensitive values
            if "SECRET" in var or "KEY" in var:
                display = f"{value[:10]}...{value[-4:]}" if len(value) > 14 else "***"
            else:
                display = value
            print(f"✅ {var}: {display}")
        else:
            print(f"❌ {var}: NOT SET")
            missing.append(f"{var} ({description})")
    
    if missing:
        print("\n⚠️  MISSING REQUIRED VARIABLES:")
        for var in missing:
            print(f"   - {var}")
        print("\nPlease add these to your .env file before continuing.")
        return False
    
    print("\n✅ All environment variables are configured!")
    return True

def test_authentication():
    """Test OAuth2 authentication"""
    print_header("🔑 STEP 2: Testing OAuth2 Authentication")
    
    try:
        manager = GraphWebhookManager()
        token = manager.get_access_token()
        
        print(f"✅ Successfully acquired access token!")
        print(f"   Token preview: {token[:20]}...")
        return True
        
    except Exception as e:
        print(f"❌ Authentication failed: {e}")
        print("\nTroubleshooting:")
        print("   1. Verify CLIENT_ID in Azure Portal")
        print("   2. Verify CLIENT_SECRET is the VALUE (not ID)")
        print("   3. Check API permissions are granted")
        return False

def list_existing_subscriptions():
    """List any existing webhook subscriptions"""
    print_header("📋 STEP 3: Checking Existing Subscriptions")
    
    try:
        manager = GraphWebhookManager()
        subscriptions = manager.list_subscriptions()
        
        if subscriptions:
            print(f"\n⚠️  Found {len(subscriptions)} existing subscription(s)")
            print("   You may want to delete old subscriptions before creating new ones.")
            
            for i, sub in enumerate(subscriptions, 1):
                print(f"\n   Subscription {i}:")
                print(f"   - ID: {sub['id']}")
                print(f"   - Resource: {sub['resource']}")
                print(f"   - Expires: {sub['expirationDateTime']}")
            
            return subscriptions
        else:
            print("✅ No existing subscriptions found")
            return []
            
    except Exception as e:
        print(f"❌ Failed to list subscriptions: {e}")
        return None

def create_subscription():
    """Create a new webhook subscription"""
    print_header("📡 STEP 4: Creating Webhook Subscription")
    
    webhook_url = os.getenv("WEBHOOK_URL")
    
    if not webhook_url:
        print("❌ WEBHOOK_URL not set in .env")
        return False
    
    if not webhook_url.startswith("https://"):
        print("⚠️  WARNING: WEBHOOK_URL must use HTTPS")
        print(f"   Current URL: {webhook_url}")
        print("   Microsoft requires HTTPS for webhooks")
        print("\nFor local development, use ngrok:")
        print("   1. Install: choco install ngrok")
        print("   2. Run: ngrok http 8001")
        print("   3. Copy HTTPS URL to WEBHOOK_URL in .env")
        return False
    
    print(f"Creating subscription for: {webhook_url}/api/webhooks/graph-notifications")
    
    try:
        manager = GraphWebhookManager()
        subscription = manager.create_subscription()
        manager.save_subscription_info()
        
        print("\n✅ Subscription created successfully!")
        print(f"   Subscription ID: {subscription['id']}")
        print(f"   Expires: {subscription['expirationDateTime']}")
        print(f"   Resource: {subscription['resource']}")
        
        return True
        
    except Exception as e:
        print(f"❌ Failed to create subscription: {e}")
        print("\nCommon issues:")
        print("   1. Webhook URL not publicly accessible")
        print("   2. FastAPI server not running")
        print("   3. Firewall blocking webhook endpoint")
        print("   4. Missing API permissions (Mail.Read)")
        return False

def verify_webhook_endpoint():
    """Test if webhook endpoint is accessible"""
    print_header("🌐 STEP 5: Verifying Webhook Endpoint")
    
    webhook_url = os.getenv("WEBHOOK_URL")
    
    if not webhook_url:
        print("❌ WEBHOOK_URL not set")
        return False
    
    import requests
    
    test_url = f"{webhook_url}/api/webhooks/subscriptions"
    
    print(f"Testing endpoint: {test_url}")
    
    try:
        response = requests.get(test_url, timeout=10)
        
        if response.status_code == 200:
            print("✅ Webhook endpoint is accessible!")
            data = response.json()
            print(f"   Active subscriptions: {data.get('count', 0)}")
            return True
        else:
            print(f"⚠️  Endpoint returned status {response.status_code}")
            return False
            
    except requests.exceptions.RequestException as e:
        print(f"❌ Cannot reach webhook endpoint: {e}")
        print("\nMake sure:")
        print("   1. FastAPI server is running (python main.py)")
        print("   2. ngrok tunnel is active (if using ngrok)")
        print("   3. WEBHOOK_URL in .env is correct")
        return False

def main():
    """Main setup flow"""
    print("============================================================")
    print("🎫 MICROSOFT GRAPH WEBHOOK SETUP")
    print("   Automated Email-to-Ticket System")
    print("============================================================")
    
    # Step 1: Check environment
    if not check_environment():
        print("\n❌ Setup cannot continue without required environment variables")
        sys.exit(1)
    
    # Step 2: Test authentication
    if not test_authentication():
        print("\n❌ Setup cannot continue without valid authentication")
        sys.exit(1)
    
    # Step 3: List existing subscriptions
    existing = list_existing_subscriptions()
    
    if existing and len(existing) > 0:
        print("\n⚠️  Delete old subscriptions? (y/n): ", end="")
        choice = input().lower()
        
        if choice == 'y':
            manager = GraphWebhookManager()
            for sub in existing:
                print(f"   Deleting {sub['id']}...")
                manager.delete_subscription(sub['id'])
            print("✅ Old subscriptions deleted")
    
    # Step 4: Verify webhook endpoint is running
    if not verify_webhook_endpoint():
        print("\n⚠️  Webhook endpoint is not accessible")
        print("   Start your FastAPI server first:")
        print("   cd backend && python main.py")
        print("\n   Then re-run this setup script")
        sys.exit(1)
    
    # Step 5: Create new subscription
    print("\n📡 Ready to create webhook subscription")
    print("   Press Enter to continue (or Ctrl+C to cancel): ", end="")
    input()
    
    if create_subscription():
        print("\n" + "="*60)
        print("🎉 SETUP COMPLETE!")
        print("="*60)
        print("\n✅ Your email automation is now active!")
        print("\nWhat happens next:")
        print("   1. New emails arrive in your inbox")
        print("   2. Microsoft sends notification to your webhook")
        print("   3. System fetches email details")
        print("   4. AI analyzes the email")
        print("   5. Ticket is created automatically")
        print("   6. Confirmation email sent to sender")
        print("\n📝 To monitor activity:")
        print("   - Check backend console logs")
        print("   - View tickets in Outlook add-in")
        print("   - Check database: backend/support_tickets.db")
        print("\n⏰ Subscription renewal:")
        print("   - Expires in 3 days")
        print("   - Run: python subscription_renewal_service.py")
        print("   - Or manually renew via API")
        print("\n🧪 Test it:")
        print("   - Send an email to your monitored inbox")
        print("   - Watch console for processing logs")
        print("   - Check for ticket creation")
        print("\n✨ Enjoy your automated ticket system!")
    else:
        print("\n❌ Setup incomplete - please fix errors above and try again")
        sys.exit(1)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n🛑 Setup cancelled by user")
        sys.exit(0)
