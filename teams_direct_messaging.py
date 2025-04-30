import requests
import json
import msal
import time
import streamlit as st

class TeamsMessenger:
    """Simplified class that uses welcome messages for notifications"""
    
    def __init__(self, client_id, client_secret, tenant_id):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = None
        
        # Create MSAL application
        self.app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=f"https://login.microsoftonline.com/{tenant_id}"
        )
    
    def authenticate(self):
        """Get access token using client credentials flow"""
        result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        
        if "access_token" in result:
            self.access_token = result["access_token"]
            return True
        else:
            return False
    
    def notify_user(self, user_id, message_text):
        """
        Send a notification by creating a new chat with welcome message
        """
        if not self.access_token and not self.authenticate():
            return False
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Create a unique topic name to avoid conflicts
        timestamp = int(time.time())
        topic = f"Timesheet Alert {timestamp}"
        
        # Create chat data with welcome message
        chat_data = {
            "chatType": "group",
            "topic": topic,
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_id}"
                }
            ],
            "welcomeMessage": {
                "content": message_text
            }
        }
        
        try:
            # Create the chat with welcome message
            url = "https://graph.microsoft.com/beta/chats"
            
            response = requests.post(url, headers=headers, json=chat_data)
            
            # Check response
            if response.status_code in [200, 201]:
                # Success!
                return True
            else:
                # Try to get error details for logging
                try:
                    error_data = response.json()
                    print(f"Error: {error_data}")
                except:
                    print(f"Error status code: {response.status_code}")
                return False
        except Exception as e:
            print(f"Exception: {str(e)}")
            return False