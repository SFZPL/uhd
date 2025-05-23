import requests
import json
import msal
import time
import streamlit as st
import logging

logger = logging.getLogger(__name__)

class TeamsMessenger:
    """Teams messenger that creates effective notification chats"""
    
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
        try:
            logger.info("Starting authentication...")
            result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                logger.info("Successfully authenticated with Microsoft Graph")
                return True
            else:
                logger.error(f"Failed to authenticate: {result.get('error_description', result)}")
                return False
        except Exception as e:
            logger.error(f"Authentication exception: {str(e)}", exc_info=True)
            return False
    
    def notify_user(self, user_id, message_text):
        """
        Send notification by creating a chat with a descriptive topic
        """
        logger.info(f"Starting notification process for user: {user_id}")
        
        if not self.access_token:
            logger.info("No access token found, attempting to authenticate...")
            if not self.authenticate():
                logger.error("Failed to get access token")
                return False
                
        # Step 1: Create the chat first
        chat_id = self._create_notification_chat(user_id, message_text)
        if not chat_id:
            logger.error("Could not create notification chat")
            return False
        
        # Step 2: Wait a moment for the chat to fully initialize
        time.sleep(2)
        
        # Step 3: Try to send a simple text message
        try:
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json"
            }
            
            # Create a simple text message (no HTML formatting)
            simple_message = f"""
            **TIMESHEET ALERT**
            
            {message_text}
            
            **Please log your missing hours in Odoo as soon as possible.**
            
            If you need assistance, please contact your manager.
            """
            
            # Prepare the request body
            message_data = {
                "body": {
                    "content": simple_message,
                    "contentType": "text"  # Use plain text instead of HTML
                }
            }
            
            # Send the message
            url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
            logger.info(f"Sending message to chat: {chat_id}")
            
            response = requests.post(url, headers=headers, json=message_data)
            
            if response.status_code in [200, 201]:
                logger.info(f"Message sent successfully to chat: {chat_id}")
                return True
            else:
                logger.error(f"Error sending message: {response.status_code} - {response.text}")
                # Even if the message fails, return True because the chat was created successfully
                # This ensures the user still gets the notification via the chat name
                return True
                
        except Exception as e:
            logger.error(f"Exception sending message: {str(e)}", exc_info=True)
            # Return True anyway since the chat was created with the notification in the topic
            return True
    
    def _create_notification_chat(self, user_id, message_text):
        """Create a chat with the notification as the topic"""
        try:
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json"
            }
            
            # Format message and clean it for topic (remove invalid characters)
            clean_message = message_text.replace(':', ' -').replace('\n', ' ')
            timestamp = time.strftime('%Y-%m-%d %H-%M')
            
            # Add email check reminder - KEEP THIS REMINDER
            email_reminder = "- Please check your email for more info"
            
            # Create topic that serves as the notification
            topic = f"{clean_message} {email_reminder} [{timestamp}]"
            
            # Ensure the topic isn't too long (200 char limit)
            if len(topic) > 200:
                # Truncate the message part while keeping the email reminder and timestamp
                max_message_length = 200 - len(email_reminder) - len(timestamp) - 5  # 5 for brackets and spaces
                clean_message = clean_message[:max_message_length-3] + "..."
                topic = f"{clean_message} {email_reminder} [{timestamp}]"
            
            # Create chat with notification in topic
            chat_data = {
                "chatType": "group",
                "topic": topic,
                "members": [
                    {
                        "@odata.type": "#microsoft.graph.aadUserConversationMember",
                        "roles": ["owner"],
                        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_id}"
                    }
                ]
            }
            
            url = "https://graph.microsoft.com/v1.0/chats"
            logger.info(f"Creating notification chat with topic: {topic}")
            
            response = requests.post(url, headers=headers, json=chat_data)
            
            if response.status_code in [200, 201]:
                chat_id = response.json().get("id")
                logger.info(f"Created notification chat: {chat_id}")
                return chat_id
            else:
                logger.error(f"Error creating chat: {response.status_code} - {response.text}")
                return None
        except Exception as e:
            logger.error(f"Exception creating chat: {str(e)}", exc_info=True)
            return None