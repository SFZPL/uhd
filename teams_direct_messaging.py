import requests
import json
import msal
import time
import streamlit as st
import logging

logger = logging.getLogger(__name__)

class TeamsMessenger:
    """Complete Teams messenger that can create chats and send messages"""
    
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
            logger.info("Successfully authenticated with Microsoft Graph")
            return True
        else:
            logger.error(f"Failed to authenticate: {result.get('error_description', result)}")
            return False
    
    def notify_user(self, user_id, message_text):
        """
        Send a notification to a user by creating a chat and sending a message
        """
        logger.info(f"Starting notification process for user: {user_id}")
        
        if not self.access_token and not self.authenticate():
            logger.error("Failed to get access token")
            return False
            
        # Step 1: Create or find existing chat
        chat_id = self._create_or_get_chat(user_id)
        if not chat_id:
            logger.error("Failed to create or get chat")
            return False
            
        logger.info(f"Got chat ID: {chat_id}")
        
        # Step 2: Send message to the chat
        result = self._send_message_to_chat(chat_id, message_text)
        logger.info(f"Message sending result: {result}")
        return result
    
    def _create_or_get_chat(self, user_id):
        """Create a new chat or get existing chat with user"""
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Create chat data for one-on-one chat
        chat_data = {
            "chatType": "oneOnOne",
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_id}"
                }
            ]
        }
        
        try:
            # Create the chat
            url = "https://graph.microsoft.com/v1.0/chats"
            
            logger.info(f"Creating chat with URL: {url}")
            logger.info(f"Chat data: {json.dumps(chat_data, indent=2)}")
            
            response = requests.post(url, headers=headers, json=chat_data)
            
            logger.info(f"Chat creation response status: {response.status_code}")
            logger.info(f"Chat creation response body: {response.text}")
            
            if response.status_code in [200, 201]:
                # Success - new chat created
                chat_id = response.json().get("id")
                logger.info(f"Successfully created/retrieved chat ID: {chat_id}")
                return chat_id
            elif response.status_code == 409:  # Conflict - chat already exists
                # Try to extract existing chat ID from error
                try:
                    error_data = response.json()
                    error_message = error_data.get("error", {}).get("message", "")
                    
                    # Extract chat ID from error message
                    import re
                    match = re.search(r"'([^']+)'", error_message)
                    if match:
                        chat_id = match.group(1)
                        logger.info(f"Found existing chat ID: {chat_id}")
                        return chat_id
                except Exception as e:
                    logger.error(f"Error parsing conflict response: {e}")
                    return None
            else:
                # Failed with non-conflict error
                logger.error(f"Error creating chat: {response.status_code} {response.text}")
                return None
        except Exception as e:
            logger.error(f"Exception creating chat: {e}", exc_info=True)
            return None
    
    def _send_message_to_chat(self, chat_id, message_text):
        """Send a message to a chat using Teamwork.Migrate.All permission"""
        if not chat_id:
            logger.error("No chat ID provided")
            return False
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Prepare message data
        message_data = {
            "body": {
                "contentType": "text",
                "content": message_text
            }
        }
        
        try:
            # Send the message
            url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
            
            logger.info(f"Sending message to URL: {url}")
            logger.info(f"Message data: {json.dumps(message_data, indent=2)}")
            
            # Adding a short delay to ensure the chat is fully created
            time.sleep(2)
            
            response = requests.post(url, headers=headers, json=message_data)
            
            logger.info(f"Message send response status: {response.status_code}")
            logger.info(f"Message send response body: {response.text}")
            
            if response.status_code in [200, 201]:
                logger.info("Message sent successfully")
                return True
            else:
                logger.error(f"Error sending message: {response.status_code} {response.text}")
                return False
        except Exception as e:
            logger.error(f"Exception sending message: {e}", exc_info=True)
            return False