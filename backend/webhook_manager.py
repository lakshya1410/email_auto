"""
Microsoft Graph API Webhook Manager
Handles subscription creation, renewal, and notification processing
"""

import os
import json
import time
import requests
from datetime import datetime, timedelta
from dotenv import load_dotenv
from typing import Optional, Dict, Any

load_dotenv()

class GraphWebhookManager:
    """Manages Microsoft Graph API webhooks for email notifications"""
    
    def __init__(self):
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self.tenant_id = os.getenv("TENANT_ID", "common")
        self.notification_url = os.getenv("WEBHOOK_URL")  # Your public HTTPS endpoint
        self.user_email = os.getenv("IMAP_EMAIL")
        
        if not all([self.client_id, self.client_secret, self.notification_url]):
            raise ValueError("Missing required environment variables: CLIENT_ID, CLIENT_SECRET, WEBHOOK_URL")
        
        self.token = None
        self.token_expires_at = None
        self.subscription_id = None
        self.subscription_expires_at = None
        
    def get_access_token(self) -> str:
        """Get OAuth2 access token using client credentials flow"""
        
        # Check if we have a valid cached token
        if self.token and self.token_expires_at and datetime.now() < self.token_expires_at:
            return self.token
        
        print("🔑 Getting new access token...")
        
        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        
        data = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials"
        }
        
        try:
            response = requests.post(token_url, data=data)
            response.raise_for_status()
            
            token_data = response.json()
            self.token = token_data["access_token"]
            expires_in = token_data.get("expires_in", 3600)
            self.token_expires_at = datetime.now() + timedelta(seconds=expires_in - 300)  # 5 min buffer
            
            print("✅ Access token acquired successfully!")
            return self.token
            
        except requests.exceptions.RequestException as e:
            print(f"❌ Failed to get access token: {e}")
            if hasattr(e.response, 'text'):
                print(f"   Response: {e.response.text}")
            raise
    
    def create_subscription(self, resource: str = "me/mailFolders/inbox/messages") -> Dict[str, Any]:
        """
        Create a webhook subscription for new emails
        
        Args:
            resource: Graph API resource to subscribe to (default: inbox messages)
            
        Returns:
            Subscription details including subscription ID
        """
        print("📡 Creating webhook subscription...")
        
        token = self.get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        # Subscription expires in 3 days (maximum for mail resources)
        expiration_time = datetime.utcnow() + timedelta(days=3)
        
        subscription_data = {
            "changeType": "created",  # Notify on new emails only
            "notificationUrl": self.notification_url,
            "resource": resource,
            "expirationDateTime": expiration_time.strftime("%Y-%m-%dT%H:%M:%S.0000000Z"),
            "clientState": "SecretClientState",  # Security token to verify notifications
            "latestSupportedTlsVersion": "v1_2"
        }
        
        try:
            response = requests.post(
                "https://graph.microsoft.com/v1.0/subscriptions",
                headers=headers,
                json=subscription_data
            )
            response.raise_for_status()
            
            subscription = response.json()
            self.subscription_id = subscription["id"]
            self.subscription_expires_at = datetime.strptime(
                subscription["expirationDateTime"], 
                "%Y-%m-%dT%H:%M:%S.%fZ"
            )
            
            print(f"✅ Subscription created successfully!")
            print(f"   Subscription ID: {self.subscription_id}")
            print(f"   Expires at: {self.subscription_expires_at}")
            
            return subscription
            
        except requests.exceptions.RequestException as e:
            print(f"❌ Failed to create subscription: {e}")
            if hasattr(e.response, 'text'):
                print(f"   Response: {e.response.text}")
            raise
    
    def renew_subscription(self, subscription_id: Optional[str] = None) -> Dict[str, Any]:
        """
        Renew an existing subscription
        
        Args:
            subscription_id: ID of subscription to renew (uses stored ID if not provided)
            
        Returns:
            Updated subscription details
        """
        sub_id = subscription_id or self.subscription_id
        if not sub_id:
            raise ValueError("No subscription ID provided or stored")
        
        print(f"🔄 Renewing subscription {sub_id}...")
        
        token = self.get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        # Extend expiration by 3 days
        new_expiration = datetime.utcnow() + timedelta(days=3)
        
        update_data = {
            "expirationDateTime": new_expiration.strftime("%Y-%m-%dT%H:%M:%S.0000000Z")
        }
        
        try:
            response = requests.patch(
                f"https://graph.microsoft.com/v1.0/subscriptions/{sub_id}",
                headers=headers,
                json=update_data
            )
            response.raise_for_status()
            
            subscription = response.json()
            self.subscription_expires_at = datetime.strptime(
                subscription["expirationDateTime"],
                "%Y-%m-%dT%H:%M:%S.%fZ"
            )
            
            print(f"✅ Subscription renewed until: {self.subscription_expires_at}")
            return subscription
            
        except requests.exceptions.RequestException as e:
            print(f"❌ Failed to renew subscription: {e}")
            if hasattr(e.response, 'text'):
                print(f"   Response: {e.response.text}")
            raise
    
    def delete_subscription(self, subscription_id: Optional[str] = None) -> bool:
        """
        Delete a webhook subscription
        
        Args:
            subscription_id: ID of subscription to delete (uses stored ID if not provided)
            
        Returns:
            True if successful
        """
        sub_id = subscription_id or self.subscription_id
        if not sub_id:
            raise ValueError("No subscription ID provided or stored")
        
        print(f"🗑️  Deleting subscription {sub_id}...")
        
        token = self.get_access_token()
        headers = {
            "Authorization": f"Bearer {token}"
        }
        
        try:
            response = requests.delete(
                f"https://graph.microsoft.com/v1.0/subscriptions/{sub_id}",
                headers=headers
            )
            response.raise_for_status()
            
            print("✅ Subscription deleted successfully!")
            self.subscription_id = None
            self.subscription_expires_at = None
            return True
            
        except requests.exceptions.RequestException as e:
            print(f"❌ Failed to delete subscription: {e}")
            if hasattr(e.response, 'text'):
                print(f"   Response: {e.response.text}")
            return False
    
    def list_subscriptions(self) -> list:
        """
        List all active subscriptions
        
        Returns:
            List of subscription objects
        """
        print("📋 Fetching active subscriptions...")
        
        token = self.get_access_token()
        headers = {
            "Authorization": f"Bearer {token}"
        }
        
        try:
            response = requests.get(
                "https://graph.microsoft.com/v1.0/subscriptions",
                headers=headers
            )
            response.raise_for_status()
            
            data = response.json()
            subscriptions = data.get("value", [])
            
            print(f"✅ Found {len(subscriptions)} active subscription(s)")
            for sub in subscriptions:
                print(f"   ID: {sub['id']}")
                print(f"   Resource: {sub['resource']}")
                print(f"   Expires: {sub['expirationDateTime']}")
                print()
            
            return subscriptions
            
        except requests.exceptions.RequestException as e:
            print(f"❌ Failed to list subscriptions: {e}")
            if hasattr(e.response, 'text'):
                print(f"   Response: {e.response.text}")
            return []
    
    def get_email_details(self, message_id: str) -> Dict[str, Any]:
        """
        Fetch full email details from Graph API
        
        Args:
            message_id: ID of the email message
            
        Returns:
            Email details including subject, body, sender
        """
        token = self.get_access_token()
        headers = {
            "Authorization": f"Bearer {token}"
        }
        
        try:
            response = requests.get(
                f"https://graph.microsoft.com/v1.0/me/messages/{message_id}",
                headers=headers,
                params={
                    "$select": "subject,from,body,receivedDateTime,sender,toRecipients"
                }
            )
            response.raise_for_status()
            
            return response.json()
            
        except requests.exceptions.RequestException as e:
            print(f"❌ Failed to get email details: {e}")
            if hasattr(e.response, 'text'):
                print(f"   Response: {e.response.text}")
            raise
    
    def validate_notification(self, validation_token: str) -> str:
        """
        Handle Microsoft's webhook validation request
        
        Args:
            validation_token: Token sent by Microsoft for validation
            
        Returns:
            The validation token (must be returned in response)
        """
        print("✅ Webhook validation request received")
        return validation_token
    
    def save_subscription_info(self, filepath: str = "subscription_info.json"):
        """Save subscription info to file for persistence"""
        if not self.subscription_id:
            print("⚠️  No subscription to save")
            return
        
        data = {
            "subscription_id": self.subscription_id,
            "expires_at": self.subscription_expires_at.isoformat() if self.subscription_expires_at else None,
            "created_at": datetime.now().isoformat()
        }
        
        with open(filepath, 'w') as f:
            json.dump(data, f, indent=2)
        
        print(f"💾 Subscription info saved to {filepath}")
    
    def load_subscription_info(self, filepath: str = "subscription_info.json") -> bool:
        """Load subscription info from file"""
        try:
            with open(filepath, 'r') as f:
                data = json.load(f)
            
            self.subscription_id = data.get("subscription_id")
            expires_str = data.get("expires_at")
            if expires_str:
                self.subscription_expires_at = datetime.fromisoformat(expires_str)
            
            print(f"📂 Subscription info loaded from {filepath}")
            return True
            
        except FileNotFoundError:
            print(f"⚠️  No subscription file found at {filepath}")
            return False
        except Exception as e:
            print(f"❌ Failed to load subscription info: {e}")
            return False


if __name__ == "__main__":
    """Test the webhook manager"""
    print("============================================================")
    print("🎫 MICROSOFT GRAPH WEBHOOK MANAGER - TEST")
    print("============================================================")
    
    manager = GraphWebhookManager()
    
    # Test token acquisition
    try:
        token = manager.get_access_token()
        print(f"✅ Token acquired: {token[:20]}...")
    except Exception as e:
        print(f"❌ Failed: {e}")
        exit(1)
    
    # List existing subscriptions
    print("\n" + "="*60)
    manager.list_subscriptions()
    
    print("\n✅ Webhook manager is ready!")
    print("   To create subscription, set WEBHOOK_URL in .env first")
