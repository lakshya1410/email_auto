# 🚀 Microsoft Graph Webhook Automation Guide

## Complete Real-Time Email-to-Ticket Automation

This guide explains how to set up **Option 1: Microsoft Graph API with Webhooks** for fully automated ticket creation from incoming emails.

---

## 📋 Table of Contents

1. [Overview](#overview)
2. [Architecture](#architecture)
3. [Prerequisites](#prerequisites)
4. [Azure Portal Setup](#azure-portal-setup)
5. [Local Development Setup](#local-development-setup)
6. [Production Deployment](#production-deployment)
7. [Testing](#testing)
8. [Troubleshooting](#troubleshooting)
9. [Maintenance](#maintenance)

---

## 🎯 Overview

### How It Works

1. **Email arrives** in your Outlook inbox
2. **Microsoft sends instant notification** to your webhook endpoint
3. **Your server receives notification** with message ID
4. **System fetches full email** from Graph API
5. **AI analyzes email** (Gemini)
6. **Ticket created automatically** in database
7. **Confirmation email sent** to sender

### Benefits

- ✅ **Real-time processing** (no polling delay)
- ✅ **Scalable** (handles multiple emails simultaneously)
- ✅ **Reliable** (Microsoft's infrastructure)
- ✅ **Secure** (OAuth2 + HTTPS)
- ✅ **Fully automated** (no manual intervention)

---

## 🏗️ Architecture

```
┌─────────────────┐
│  Outlook Inbox  │
└────────┬────────┘
         │ New Email
         ▼
┌─────────────────────┐
│  Microsoft Graph    │
│  Notification API   │
└────────┬────────────┘
         │ Webhook POST
         │ (with message ID)
         ▼
┌─────────────────────────────┐
│  Your FastAPI Server        │
│  (backend/main.py)          │
│                             │
│  /api/webhooks/            │
│    graph-notifications      │
└────────┬────────────────────┘
         │
         ▼
┌─────────────────────────────┐
│  Email Processing           │
│  1. Fetch email details     │
│  2. AI analysis (Gemini)    │
│  3. Create ticket           │
│  4. Send confirmation       │
└─────────────────────────────┘
```

---

## ✅ Prerequisites

### Required Accounts
- [x] **Microsoft Account** (Outlook/Hotmail/365)
- [x] **Azure Account** (free tier works)
- [x] **Google Account** (for Gemini API)

### Required Software
- [x] **Python 3.8+**
- [x] **ngrok** (for local development)
- [x] **PowerShell** or **Terminal**

### Required Skills
- Basic command line usage
- Azure Portal navigation
- Environment variable configuration

---

## 🔐 Azure Portal Setup

### Step 1: Register Azure Application

1. **Go to Azure Portal**: https://portal.azure.com
2. **Navigate to**: Azure Active Directory → App registrations
3. **Click**: "+ New registration"
4. **Fill in**:
   ```
   Name: Email Ticket Automation
   Account types: ✅ Personal Microsoft accounts only
   Redirect URI: (Leave blank)
   ```
5. **Click**: "Register"

### Step 2: Get Application Credentials

1. **Copy Application (client) ID**:
   ```
   Example: 12345678-1234-1234-1234-123456789012
   ```
   Save this - you'll need it for `.env`

2. **Copy Directory (tenant) ID**:
   ```
   Example: 87654321-4321-4321-4321-210987654321
   ```
   Or use `common` for personal accounts

### Step 3: Create Client Secret

1. **Go to**: Certificates & secrets → Client secrets
2. **Click**: "+ New client secret"
3. **Description**: `EmailAutomation`
4. **Expires**: 24 months
5. **Click**: "Add"
6. **⚠️ IMPORTANT**: Copy the **VALUE** immediately (not the ID)!
   ```
   Example: abc~def1234567890abcdefghijklmnopqrstuvwxyz
   ```
   You **cannot** see this again!

### Step 4: Configure API Permissions

1. **Go to**: API permissions
2. **Click**: "+ Add a permission"
3. **Choose**: Microsoft Graph
4. **Choose**: **Application permissions** (NOT Delegated!)
5. **Add these permissions**:
   - ✅ `Mail.Read` - Read mail in all mailboxes
   - ✅ `Mail.ReadWrite` - Read and write mail in all mailboxes
6. **Click**: "Grant admin consent for [Directory]"
7. **Verify**: All permissions show green checkmarks ✅

**⚠️ CRITICAL**: Without admin consent, authentication will fail!

---

## 💻 Local Development Setup

### Step 1: Update .env File

Add these variables to `backend/.env`:

```env
# Azure App Registration
CLIENT_ID=your-client-id-from-azure
CLIENT_SECRET=your-client-secret-value-from-azure
TENANT_ID=common

# Webhook Configuration (will be set after ngrok)
WEBHOOK_URL=https://your-ngrok-url.ngrok.io

# Email to monitor
IMAP_EMAIL=your-email@outlook.com

# Keep existing variables
GEMINI_API_KEY=your-gemini-key
SMTP_EMAIL=your-email@outlook.com
SMTP_PASSWORD=your-app-password
```

### Step 2: Install ngrok (Public URL Tunnel)

**Why ngrok?** Microsoft webhooks require HTTPS. ngrok creates a secure tunnel from internet to your localhost.

#### Install ngrok:

**Windows (PowerShell as Admin)**:
```powershell
choco install ngrok
```

Or download from: https://ngrok.com/download

#### Authenticate ngrok:
1. Sign up at https://dashboard.ngrok.com/signup
2. Get your auth token
3. Run:
   ```powershell
   ngrok config add-authtoken YOUR_AUTH_TOKEN
   ```

### Step 3: Start ngrok Tunnel

```powershell
ngrok http 8001
```

**Output**:
```
Session Status    online
Region            United States (us)
Forwarding        https://abc123.ngrok.io -> http://localhost:8001
```

**⚠️ COPY** the `https://abc123.ngrok.io` URL!

### Step 4: Update WEBHOOK_URL in .env

```env
WEBHOOK_URL=https://abc123.ngrok.io
```

**Note**: ngrok URLs change every time you restart. Update this each time!

### Step 5: Install Python Dependencies

```powershell
cd backend
pip install requests schedule python-dotenv
```

### Step 6: Start FastAPI Server

```powershell
cd backend
python main.py
```

**Expected output**:
```
✅ Gemini API configured successfully
INFO:     Started server process [12345]
INFO:     Uvicorn running on http://0.0.0.0:8001
```

### Step 7: Run Setup Script

**In a NEW terminal**:

```powershell
cd backend
python setup_webhooks.py
```

**The script will**:
1. ✅ Check environment variables
2. ✅ Test OAuth2 authentication
3. ✅ List existing subscriptions
4. ✅ Verify webhook endpoint
5. ✅ Create new subscription

**Expected output**:
```
============================================================
🎫 MICROSOFT GRAPH WEBHOOK SETUP
============================================================

🔍 STEP 1: Checking Environment Variables
✅ CLIENT_ID: 12345678...
✅ CLIENT_SECRET: abc~def...
✅ WEBHOOK_URL: https://abc123.ngrok.io
✅ All environment variables configured!

🔑 STEP 2: Testing OAuth2 Authentication
✅ Successfully acquired access token!

📋 STEP 3: Checking Existing Subscriptions
✅ No existing subscriptions found

🌐 STEP 5: Verifying Webhook Endpoint
✅ Webhook endpoint is accessible!

📡 STEP 4: Creating Webhook Subscription
✅ Subscription created successfully!
   Subscription ID: abc-123-def
   Expires: 2025-10-30T10:30:00Z

🎉 SETUP COMPLETE!
```

---

## 🧪 Testing

### Test 1: Send Email

1. Send an email to your monitored inbox (`IMAP_EMAIL`)
2. Watch backend console for logs:

```
📬 Received 1 notification(s) from Microsoft Graph

📧 Processing email notification: AAMkAD...
📥 Fetching email details from Graph API...
   From: John Doe <john@example.com>
   Subject: Need help with account
🎫 Generated ticket number: TKT-000001
🤖 Analyzing email with Gemini AI...
   Category: Support
   Priority: Medium
✅ Ticket TKT-000001 created successfully!
📧 Sending confirmation email...
✅ Confirmation email sent!
✨ Processing complete for ticket TKT-000001
```

### Test 2: Verify Ticket Created

1. Open Outlook add-in
2. Click **Tickets** tab
3. You should see new ticket: **TKT-000001**

### Test 3: Check Confirmation Email

1. Check sender's email inbox
2. Should receive: "Your Support Ticket #TKT-000001"

---

## 🌍 Production Deployment

For production, you need a **real public server** (not ngrok).

### Option A: Azure App Service

1. **Create App Service**:
   ```bash
   az webapp create --name email-ticket-api --resource-group MyResourceGroup
   ```

2. **Set environment variables** in Azure Portal:
   - Configuration → Application settings
   - Add all `.env` variables

3. **Update WEBHOOK_URL**:
   ```env
   WEBHOOK_URL=https://email-ticket-api.azurewebsites.net
   ```

4. **Deploy**:
   ```bash
   az webapp up --name email-ticket-api
   ```

### Option B: AWS EC2

1. **Launch Ubuntu instance**
2. **Install dependencies**:
   ```bash
   sudo apt update
   sudo apt install python3-pip nginx certbot
   ```

3. **Set up HTTPS** with Let's Encrypt:
   ```bash
   sudo certbot --nginx -d yourdomain.com
   ```

4. **Configure nginx reverse proxy** to port 8001
5. **Update WEBHOOK_URL** to your domain

### Option C: DigitalOcean / Heroku / Railway

Similar process - ensure:
- ✅ HTTPS enabled
- ✅ Port 8001 accessible
- ✅ Environment variables set
- ✅ WEBHOOK_URL points to your domain

---

## 🔄 Maintenance

### Subscription Renewal

Subscriptions expire every **3 days**. Options to renew:

#### Option 1: Automatic Renewal Service

Run this in the background:

```powershell
cd backend
python subscription_renewal_service.py
```

This checks every 12 hours and auto-renews if expiring within 24 hours.

#### Option 2: Manual Renewal

```powershell
curl -X POST http://localhost:8001/api/webhooks/subscriptions/renew
```

#### Option 3: Scheduled Task (Windows)

Create a Windows Task:
1. Open Task Scheduler
2. Create Basic Task → "Renew Webhook Subscription"
3. Trigger: Daily at 2 AM
4. Action: Start Program → `python`
5. Arguments: `C:\path\to\backend\subscription_renewal_service.py`

---

## 🔧 Troubleshooting

### Issue: "Failed to create subscription: invalid_client"

**Solution**:
- Verify `CLIENT_ID` in `.env` matches Azure Portal
- Verify `CLIENT_SECRET` is the **VALUE** (not ID)
- Regenerate secret if unsure

### Issue: "Insufficient privileges to complete the operation"

**Solution**:
- Go to Azure Portal → App registrations → Your app → API permissions
- Ensure **Application permissions** (not Delegated)
- Click "Grant admin consent" again
- Wait 5-10 minutes for propagation

### Issue: "Subscription validation failed"

**Solution**:
- Ensure FastAPI server is running
- Ensure ngrok tunnel is active
- Verify `WEBHOOK_URL` is correct and accessible
- Check firewall/antivirus not blocking webhook endpoint

### Issue: "Emails not triggering notifications"

**Solution**:
- Check subscription is active: `GET /api/webhooks/subscriptions`
- Verify email is sent to correct inbox (`IMAP_EMAIL`)
- Check subscription hasn't expired (3 day limit)
- Look for errors in backend console logs

### Issue: "ngrok URL keeps changing"

**Solutions**:
- **Free plan**: Restart setup each time
- **Paid plan** ($5/month): Get permanent URL
- **Production**: Deploy to real server

---

## 📊 Monitoring

### Check Subscription Status

```powershell
# List all subscriptions
curl http://localhost:8001/api/webhooks/subscriptions

# Response:
{
  "count": 1,
  "subscriptions": [
    {
      "id": "abc-123-def",
      "resource": "me/mailFolders/inbox/messages",
      "expirationDateTime": "2025-10-30T10:30:00Z"
    }
  ]
}
```

### View Backend Logs

All webhook activity is logged to console:
- 📬 Notification received
- 📧 Email being processed
- 🎫 Ticket created
- 📧 Confirmation sent

### Check Database

```bash
sqlite3 backend/support_tickets.db
SELECT ticket_number, status, subject FROM support_tickets ORDER BY created_at DESC LIMIT 10;
```

---

## 🎯 Best Practices

### Security

1. **Keep CLIENT_SECRET secure** - never commit to git
2. **Use HTTPS only** - Microsoft requires it
3. **Validate clientState** - prevents spoofed notifications
4. **Rotate secrets** - every 6-12 months

### Reliability

1. **Monitor subscription expiry** - use renewal service
2. **Log all webhook events** - for debugging
3. **Handle duplicates** - email_hash prevents duplicate tickets
4. **Graceful error handling** - don't crash on malformed notifications

### Performance

1. **Process async** - webhook returns immediately, processes in background
2. **Use thread pool** - handle multiple emails simultaneously
3. **Cache tokens** - don't fetch new token for every request
4. **Rate limiting** - respect Graph API limits (10,000 requests/10 min)

---

## 📚 Additional Resources

- **Microsoft Graph Webhooks**: https://docs.microsoft.com/en-us/graph/webhooks
- **Graph API Explorer**: https://developer.microsoft.com/en-us/graph/graph-explorer
- **Azure Portal**: https://portal.azure.com
- **ngrok Documentation**: https://ngrok.com/docs
- **FastAPI Webhooks**: https://fastapi.tiangolo.com/advanced/events/

---

## ✅ Quick Start Checklist

- [ ] Azure app registered
- [ ] Client ID and secret saved
- [ ] API permissions granted with admin consent
- [ ] .env file updated with credentials
- [ ] ngrok installed and authenticated
- [ ] ngrok tunnel running (`ngrok http 8001`)
- [ ] WEBHOOK_URL set to ngrok URL
- [ ] Dependencies installed (`pip install -r requirements.txt`)
- [ ] FastAPI server running (`python main.py`)
- [ ] Setup script executed (`python setup_webhooks.py`)
- [ ] Webhook subscription created successfully
- [ ] Test email sent and ticket created
- [ ] Renewal service running (optional for production)

---

## 🎉 Success!

Once setup is complete, your system will:
- ✅ Automatically create tickets for ALL incoming emails
- ✅ Process emails in real-time (no delay)
- ✅ Send instant confirmation to senders
- ✅ Display tickets in Outlook add-in
- ✅ Track analytics in dashboard

**Enjoy your fully automated support ticket system!** 🚀

---

## 💬 Need Help?

If you encounter issues:
1. Check troubleshooting section above
2. Verify all checklist items
3. Review backend console logs
4. Test with Graph API Explorer
5. Check Azure Portal for permission issues

**Common mistakes**:
- Using Delegated instead of Application permissions
- Not granting admin consent
- CLIENT_SECRET is the ID instead of VALUE
- ngrok not running or URL incorrect
- FastAPI server not started before running setup
