import requests
import os
import streamlit as st
import json
import logging
import msal
import traceback

# Configure logging
logger = logging.getLogger(__name__)

class TeamsDirectMessaging:
    """Simplified class to handle Microsoft Teams direct messaging via Graph API"""
    
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
        """
        Get user ID by email using the filter endpoint
        This works with application permissions if you have User.Read.All
        
        If your app doesn't have this permission, 
        you'll need to manually map emails to IDs
        """
        if not self.access_token:
            if not self.authenticate():
                logger.error("Authentication failed")
                return None
                
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            # Try to find the user by email
            url = f"https://graph.microsoft.com/v1.0/users?$filter=mail eq '{email}' or userPrincipalName eq '{email}'"
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                data = response.json()
                if "value" in data and len(data["value"]) > 0:
                    user_id = data["value"][0]["id"]
                    logger.info(f"Found user ID for {email}: {user_id}")
                    return user_id
                else:
                    logger.warning(f"No user found with email {email}")
                    return None
            else:
                logger.error(f"Failed to get user by email: {response.status_code} {response.text}")
                return None
        except Exception as e:
            logger.error(f"Error in get_user_id_by_email: {e}")
            return None
            
    def create_chat(self, user_id):
        """
        Create a chat with a user
        This uses the Chat.Create application permission
        """
        if not self.access_token:
            if not self.authenticate():
                return None
                
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # For application permissions, we need to create a group chat
        # because one-on-one chats require delegated permissions
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
                return chat_id
            elif response.status_code == 409:  # Conflict - chat may already exist
                try:
                    # Try to get the existing chat ID from the response
                    error_data = response.json()
                    # Often the chat ID is included in the error response
                    error_message = error_data.get("error", {}).get("message", "")
                    
                    # Try to extract the chat ID from the error message
                    # Error messages might contain the ID in various formats
                    if "existing chat" in error_message.lower() and "'" in error_message:
                        import re
                        match = re.search(r"'([^']+)'", error_message)
                        if match:
                            chat_id = match.group(1)
                            logger.info(f"Extracted existing chat ID from error: {chat_id}")
                            return chat_id
                except Exception as e:
                    logger.error(f"Error parsing conflict response: {e}")
                
                # If we can't extract ID, fall back to searching for chats
                return self._find_existing_chat(user_id)
            else:
                logger.error(f"Failed to create chat: {response.status_code} {response.text}")
                # Try alternative approach as a fallback
                return self._find_existing_chat(user_id)
        except Exception as e:
            logger.error(f"Error in create_chat: {e}")
            return None
            
    def _find_existing_chat(self, user_id):
        """
        Try to find an existing chat with the user
        This is a fallback if direct creation fails
        """
        if not self.access_token:
            if not self.authenticate():
                return None
                
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            # Try to list all chats
            url = "https://graph.microsoft.com/v1.0/chats"
            logger.info("Attempting to list all chats to find existing chat")
            
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                data = response.json()
                if "value" in data:
                    chats = data["value"]
                    logger.info(f"Found {len(chats)} chats, looking for one with user {user_id}")
                    
                    # For each chat, try to find one that includes our target user
                    for chat in chats:
                        chat_id = chat.get("id")
                        
                        # Get members of this chat
                        members_url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/members"
                        members_response = requests.get(members_url, headers=headers)
                        
                        if members_response.status_code == 200:
                            members_data = members_response.json()
                            if "value" in members_data:
                                members = members_data["value"]
                                
                                # Check if our target user is a member
                                for member in members:
                                    member_id = member.get("userId")
                                    if member_id == user_id:
                                        logger.info(f"Found existing chat with user: {chat_id}")
                                        return chat_id
            
            # If we get here, we couldn't find an existing chat
            logger.warning(f"Could not find existing chat with user {user_id}")
            
            # Create a new chat using alternative approach
            return self._create_chat_alternative(user_id)
        except Exception as e:
            logger.error(f"Error in _find_existing_chat: {e}")
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
            
            if response.status_code in [200, 201]:
                data = response.json()
                chat_id = data.get("id")
                logger.info(f"Successfully created chat with alternative approach: {chat_id}")
                return chat_id
            else:
                logger.error(f"Alternative chat creation failed: {response.status_code} {response.text}")
                return None
        except Exception as e:
            logger.error(f"Error in _create_chat_alternative: {e}")
            return None
            
    def send_direct_message(self, chat_id, message_content):
        """
        Send a message to a chat
        This uses the ChatMessage.Send application permission
        """
        if not self.access_token:
            if not self.authenticate():
                return False
                
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        message_data = {
            "body": {
                "content": message_content,
                "contentType": "html"
            }
        }
        
        try:
            url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
            logger.info(f"Sending message to chat ID: {chat_id}")
            
            response = requests.post(url, headers=headers, json=message_data)
            
            if response.status_code in [200, 201]:
                logger.info(f"Successfully sent message to chat {chat_id}")
                return True
            else:
                logger.error(f"Failed to send message: {response.status_code} {response.text}")
                return False
        except Exception as e:
            logger.error(f"Error in send_direct_message: {e}")
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