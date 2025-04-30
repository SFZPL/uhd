import requests
import os
import streamlit as st
import json
import logging
import msal
import traceback
import time

# Configure logging
logger = logging.getLogger(__name__)

class TeamsDirectMessaging:
    """Class to handle Microsoft Teams direct messaging via Graph API using application permissions"""
    
    def __init__(self, client_id, client_secret, tenant_id, access_token=None):
        """Initialize with Azure AD credentials"""
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = access_token
        
        # Log initialization (mask sensitive parts)
        client_id_masked = client_id[:5] + "..." + client_id[-5:] if client_id and len(client_id) > 10 else "None"
        tenant_id_masked = tenant_id[:5] + "..." + tenant_id[-5:] if tenant_id and len(tenant_id) > 10 else "None"
        logger.info(f"Initializing TeamsDirectMessaging with client_id: {client_id_masked}, tenant: {tenant_id_masked}")
        
        # Create MSAL application
        try:
            self.app = msal.ConfidentialClientApplication(
                client_id=client_id,
                client_credential=client_secret,
                authority=f"https://login.microsoftonline.com/{tenant_id}"
            )
            logger.info("MSAL application initialized successfully")
        except Exception as e:
            logger.error(f"Error initializing MSAL application: {e}")
            logger.error(traceback.format_exc())
            self.app = None
    
    def authenticate(self):
        """
        Authenticate using client credentials (app permissions) flow
        Returns True if authentication is successful, False otherwise
        """
        if self.access_token:
            logger.info("Using existing access token")
            return True
            
        try:
            # Acquire token using client credentials (app permissions)
            logger.info("Acquiring token using client credentials flow")
            result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                token_masked = f"{self.access_token[:10]}...{self.access_token[-10:]}" if len(self.access_token) > 20 else "short_token"
                logger.info(f"Successfully acquired token: {token_masked}")
                return True
            else:
                error = result.get("error", "unknown")
                error_desc = result.get("error_description", "No description")
                logger.error(f"Authentication failed: {error} - {error_desc}")
                return False
        except Exception as e:
            logger.error(f"Error in authentication: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def get_user_id_by_email(self, email):
        """Get Teams user ID from email address"""
        if not self.access_token:
            logger.info("No access token found, authenticating first")
            if not self.authenticate():
                logger.error("Authentication failed before get_user_id_by_email")
                return None
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            url = (
                "https://graph.microsoft.com/v1.0/users"
                f"?$filter=mail eq '{email}' or userPrincipalName eq '{email}'"
            )
            logger.info(f"Making request to: {url}")
            
            response = requests.get(url, headers=headers)
            
            # Log status code and partial response
            logger.info(f"Response status: {response.status_code}")
            if response.status_code != 200:
                logger.error(f"Response text: {response.text[:200]}...")
            
            if response.status_code == 200:
                user_data = response.json()
                if "value" in user_data and len(user_data["value"]) > 0:
                    user_id = user_data["value"][0].get("id")
                    logger.info(f"Found user ID: {user_id[:5]}...{user_id[-5:]}" if user_id and len(user_id) > 10 else f"Found user ID: {user_id}")
                    return user_id
                else:
                    logger.error(f"No users found with email {email}")
                    return None
            else:
                logger.error(f"Failed to get user ID for {email}. Status: {response.status_code}, Response: {response.text}")
                return None
        except Exception as e:
            logger.error(f"Error getting user ID: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def create_chat(self, user_id):
        """
        Create a chat with a user using application permissions
        Returns chat_id or None
        """
        if not self.access_token:
            logger.info("No access token found, authenticating first")
            if not self.authenticate():
                logger.error("Authentication failed before create_chat")
                return None
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Using application permissions we need to create a group chat
        chat_data = {
            "chatType": "group",
            "topic": "Missing Timesheet Notification",
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_id}"
                }
            ]
        }
        
        try:
            url = "https://graph.microsoft.com/v1.0/chats"
            logger.info(f"Creating chat with user ID: {user_id}")
            
            response = requests.post(url, headers=headers, json=chat_data)
            
            # Log full response for debugging
            logger.info(f"Create chat response status: {response.status_code}")
            logger.info(f"Create chat response: {response.text[:500]}")
            
            if response.status_code in [200, 201]:
                data = response.json()
                chat_id = data.get("id")
                logger.info(f"Successfully created chat with ID: {chat_id}")
                
                # Wait a moment to let the chat be fully created before sending messages
                time.sleep(2)
                return chat_id
            elif response.status_code == 409:  # Conflict - chat may already exist
                try:
                    # Try to get the existing chat ID from the response
                    error_data = response.json()
                    error_message = error_data.get("error", {}).get("message", "")
                    
                    # Try to extract the chat ID from the error message
                    if "existing chat" in error_message.lower() and "'" in error_message:
                        import re
                        match = re.search(r"'([^']+)'", error_message)
                        if match:
                            chat_id = match.group(1)
                            logger.info(f"Extracted existing chat ID from error: {chat_id}")
                            return chat_id
                except Exception as e:
                    logger.error(f"Error parsing conflict response: {e}")
                
                # Fall back to alternative method
                return self._create_chat_alternative(user_id)
            else:
                logger.error(f"Failed to create chat: {response.status_code} {response.text}")
                # Try alternative approach as a fallback
                return self._create_chat_alternative(user_id)
        except Exception as e:
            logger.error(f"Error in create_chat: {e}")
            logger.error(traceback.format_exc())
            return self._create_chat_alternative(user_id)

    def _create_chat_alternative(self, user_id):
        """
        Alternative approach to create a chat
        This tries a different format that sometimes works when the standard one fails
        """
        if not self.access_token:
            if not self.authenticate():
                return None
                
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Alternative format for chat creation
        chat_data = {
            "chatType": "group",
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_id}"
                }
            ]
        }
        
        try:
            url = "https://graph.microsoft.com/v1.0/chats"
            logger.info(f"Trying alternative chat creation for user ID: {user_id}")
            
            response = requests.post(url, headers=headers, json=chat_data)
            
            # Log full response for debugging
            logger.info(f"Alternative create chat response status: {response.status_code}")
            logger.info(f"Alternative create chat response: {response.text[:500]}")
            
            if response.status_code in [200, 201]:
                data = response.json()
                chat_id = data.get("id")
                logger.info(f"Successfully created chat with alternative approach: {chat_id}")
                
                # Wait a moment to let the chat be fully created before sending messages
                time.sleep(2)
                return chat_id
            else:
                logger.error(f"Alternative chat creation failed: {response.status_code} {response.text}")
                return None
        except Exception as e:
            logger.error(f"Error in _create_chat_alternative: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def send_direct_message(self, chat_id, message_content):
        """
        Send a message to a chat with improved error handling and fallbacks
        """
        if not self.access_token:
            if not self.authenticate():
                return False
                
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # First try: Simple text message as a test
        simple_message = {
            "body": {
                "content": "Notification from Missing Timesheet Reporter",
                "contentType": "text"
            }
        }
        
        try:
            url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
            logger.info(f"Testing message sending to chat ID: {chat_id}")
            
            # First send a simple test message
            simple_response = requests.post(url, headers=headers, json=simple_message)
            logger.info(f"Simple message response: Status {simple_response.status_code}")
            
            if simple_response.status_code in [200, 201]:
                logger.info("Simple test message sent successfully, now sending full message")
                
                # Wait a moment before sending the main message
                time.sleep(1)
                
                # Now try sending the full HTML message in parts to avoid potential size issues
                # First send the header
                header = f"""
                <div>
                <h2>Missing Timesheet Notification</h2>
                <p>Hi there,</p>
                <p>This is a notification about missing timesheet entries.</p>
                </div>
                """
                
                header_message = {
                    "body": {
                        "content": header,
                        "contentType": "html"
                    }
                }
                
                header_response = requests.post(url, headers=headers, json=header_message)
                logger.info(f"Header message response: Status {header_response.status_code}")
                
                # Now send the content as plain text
                content_text = "Please check your assigned tasks and log your hours as soon as possible."
                
                content_message = {
                    "body": {
                        "content": content_text,
                        "contentType": "text"
                    }
                }
                
                content_response = requests.post(url, headers=headers, json=content_message)
                logger.info(f"Content message response: Status {content_response.status_code}")
                
                # Success if at least the simple message was sent
                return True
            else:
                logger.error(f"Failed to send even simple message: {simple_response.status_code} {simple_response.text}")
                
                # Final fallback - try with beta endpoint
                beta_url = f"https://graph.microsoft.com/beta/chats/{chat_id}/messages"
                beta_response = requests.post(beta_url, headers=headers, json=simple_message)
                
                if beta_response.status_code in [200, 201]:
                    logger.info("Successfully sent message using beta endpoint")
                    return True
                else:
                    logger.error(f"Beta endpoint also failed: {beta_response.status_code} {beta_response.text}")
                    return False
        except Exception as e:
            logger.error(f"Error in send_direct_message: {e}", exc_info=True)
            return False
            
    def check_authentication(self):
        """
        Simple test to verify if authentication works
        Returns a dictionary with information about the token
        """
        if not self.access_token:
            if not self.authenticate():
                return {"status": "error", "message": "Authentication failed"}
                
        # Try to decode the token to get some info about it
        try:
            import base64
            import json
            
            # The token is in format: header.payload.signature
            # We want to decode the payload part
            token_parts = self.access_token.split('.')
            if len(token_parts) >= 2:
                # Get the payload part and decode it
                payload = token_parts[1]
                # Add padding if needed
                payload += '=' * (4 - len(payload) % 4) if len(payload) % 4 != 0 else ''
                
                try:
                    decoded = base64.b64decode(payload)
                    token_data = json.loads(decoded)
                    
                    # Extract useful information
                    return {
                        "status": "success",
                        "app_id": token_data.get("appid"),
                        "audience": token_data.get("aud"),
                        "tenant_id": token_data.get("tid"),
                        "roles": token_data.get("roles", []),
                        "scopes": token_data.get("scp", "")
                    }
                except Exception as e:
                    return {"status": "partial", "message": f"Token obtained but could not decode: {str(e)}"}
            else:
                return {"status": "partial", "message": "Token obtained but in unexpected format"}
        except Exception as e:
            return {"status": "partial", "message": f"Token obtained but error analyzing it: {str(e)}"}