import requests
import json
import logging
import msal
import time
import traceback

# Configure logging
logger = logging.getLogger(__name__)

class TeamsMessenger:
    """A simplified class focused solely on creating chats and posting messages"""
    
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
            logger.error(f"Authentication failed: {result.get('error')}")
            return False
    
    def notify_user(self, user_id, message_text):
        """
        Send a notification to a user
        1. First create a chat without a welcome message 
        2. Then explicitly send a message to that chat
        """
        if not self.access_token and not self.authenticate():
            logger.error("Failed to authenticate")
            return False
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # CREATE CHAT FIRST
        
        # Create a unique chat topic with timestamp
        timestamp = int(time.time())
        chat_topic = f"Missing Timesheet Alert {timestamp}"
        
        # Create a chat WITHOUT a welcome message
        chat_data = {
            "chatType": "group",
            "topic": chat_topic,
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_id}"
                },
                {
                    "@odata.type": "#microsoft.graph.aadServiceConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": "https://graph.microsoft.com/v1.0/servicePrincipals/3cb693da-c503-4a0c-afba-0408f28d77b6"
                }
            ]
        }

        
        try:
            # Create the chat 
            url = "https://graph.microsoft.com/v1.0/chats"  # Using beta endpoint for better compatibility
            logger.info(f"Creating chat for user ID: {user_id}")
            
            response = requests.post(url, headers=headers, json=chat_data)
            logger.info(f"Create chat response: {response.status_code}")
            
            chat_id = None
            
            if response.status_code in [200, 201]:
                chat_info = response.json()
                chat_id = chat_info.get("id")
                logger.info(f"Successfully created chat with ID: {chat_id}")
            elif response.status_code == 409:  # Chat already exists
                logger.info("Chat already exists, trying to extract ID from error message")
                try:
                    error_data = response.json()
                    error_message = error_data.get("error", {}).get("message", "")
                    
                    # Try to extract chat ID from error message
                    if "'" in error_message:
                        import re
                        match = re.search(r"'([^']+)'", error_message)
                        if match:
                            chat_id = match.group(1)
                            logger.info(f"Extracted existing chat ID from error: {chat_id}")
                except Exception as e:
                    logger.error(f"Error parsing conflict response: {e}")
            else:
                logger.error(f"Failed to create chat: {response.status_code} {response.text}")
                return False
            
            # If no chat ID, return failure
            if not chat_id:
                logger.error("Failed to get a valid chat ID")
                return False
                
            # SEND MESSAGE TO CHAT
            
            # Delay to ensure chat is ready (important!)
            logger.info("Waiting 8 seconds for chat to fully initialize...")
            time.sleep(8)  # Increased delay for better reliability
            
            # Prepare a simple message
            message_data = {
                "body": {
                    "contentType": "text",
                    "content": message_text
                }
            }
            
            # Send message with beta endpoint
            message_url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
            
            # Try a few times with increasing delays
            for attempt in range(1, 4):
                message_response = requests.post(message_url, headers=headers, json=message_data)
                logger.info(f"Message response (attempt {attempt}): {message_response.status_code}")
                
                if message_response.status_code in [200, 201]:
                    logger.info("Message sent successfully!")
                    return True
                else:
                    logger.warning(f"Message sending failed on attempt {attempt}: {message_response.text}")
                    
                    # Only add extra delay if this isn't the last attempt
                    if attempt < 3:
                        wait_time = attempt * 3  # 3, 6 seconds
                        logger.info(f"Waiting {wait_time} seconds before retrying...")
                        time.sleep(wait_time)
            
            # If we get here, all attempts failed
            logger.error("All message sending attempts failed")
            
            # Try a final fallback
            return self._try_alternative_message_methods(chat_id, message_text)
                
        except Exception as e:
            logger.error(f"Error in notify_user: {e}", exc_info=True)
            return False
    
    def _try_alternative_message_methods(self, chat_id, message_text):
        """Try alternative message sending methods"""
        if not self.access_token and not self.authenticate():
            return False
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            # 1. Try using a different message format
            chat_message = {
                "body": {
                    "contentType": "html",
                    "content": f"<p>{message_text}</p>"
                }
            }
            
            message_url = f"https://graph.microsoft.com/beta/chats/{chat_id}/messages"
            
            html_response = requests.post(message_url, headers=headers, json=chat_message)
            
            if html_response.status_code in [200, 201]:
                logger.info("Message sent successfully using HTML format")
                return True
                
            logger.warning(f"HTML message format failed: {html_response.status_code}")
            
            # 2. Try using a very simple plain message
            simple_message = {
                "body": {
                    "content": "Missing Timesheet Alert: Please log your hours."
                }
            }
            
            simple_response = requests.post(message_url, headers=headers, json=simple_message)
            
            if simple_response.status_code in [200, 201]:
                logger.info("Simple message sent successfully")
                return True
                
            logger.warning(f"Simple message failed: {simple_response.status_code}")
            
            # All methods failed
            return False
            
        except Exception as e:
            logger.error(f"Error in alternative message methods: {e}")
            return False