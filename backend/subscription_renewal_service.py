"""
Subscription Renewal Service
Automatically renews Microsoft Graph webhook subscriptions before they expire
"""

import time
import schedule
from datetime import datetime, timedelta
from webhook_manager import GraphWebhookManager
import signal
import sys

class SubscriptionRenewalService:
    """Background service that renews webhook subscriptions automatically"""
    
    def __init__(self):
        self.manager = GraphWebhookManager()
        self.running = False
        
    def renew_if_needed(self):
        """Check and renew subscription if expiring soon"""
        try:
            # Load saved subscription info
            loaded = self.manager.load_subscription_info()
            
            if not loaded or not self.manager.subscription_expires_at:
                print("⚠️  No active subscription found")
                return
            
            # Check if expiring within 24 hours
            time_until_expiry = self.manager.subscription_expires_at - datetime.now()
            hours_until_expiry = time_until_expiry.total_seconds() / 3600
            
            print(f"🕐 Subscription expires in {hours_until_expiry:.1f} hours")
            
            if hours_until_expiry < 24:
                print("🔄 Subscription expiring soon - renewing now...")
                self.manager.renew_subscription()
                self.manager.save_subscription_info()
                print("✅ Subscription renewed successfully!")
            else:
                print("✅ Subscription still valid")
                
        except Exception as e:
            print(f"❌ Error checking/renewing subscription: {e}")
    
    def start(self):
        """Start the renewal service"""
        print("============================================================")
        print("🔄 SUBSCRIPTION RENEWAL SERVICE")
        print("============================================================")
        print("   Checking subscriptions every 12 hours")
        print("   Auto-renewing if expiring within 24 hours")
        print("   Press Ctrl+C to stop")
        print("============================================================\n")
        
        # Set up signal handler for graceful shutdown
        signal.signal(signal.SIGINT, self._signal_handler)
        signal.signal(signal.SIGTERM, self._signal_handler)
        
        # Schedule renewal check every 12 hours
        schedule.every(12).hours.do(self.renew_if_needed)
        
        # Run immediately on start
        self.renew_if_needed()
        
        self.running = True
        
        # Main loop
        while self.running:
            schedule.run_pending()
            time.sleep(60)  # Check every minute
    
    def _signal_handler(self, signum, frame):
        """Handle shutdown signals gracefully"""
        print("\n\n🛑 Shutdown signal received")
        print("   Stopping renewal service...")
        self.running = False
        sys.exit(0)


if __name__ == "__main__":
    service = SubscriptionRenewalService()
    service.start()
