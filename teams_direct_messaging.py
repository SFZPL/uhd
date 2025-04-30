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
        
        # First check if chat already exists by listing chats and filtering
        try:
            # Using application permissions we need to create a group chat
            # with at least one other member (we'll use the bot's own ID)
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
            
            logger.info(f"Creating chat with user ID: {user_id}")
            
            response = requests.post(
                "https://graph.microsoft.com/v1.0/chats",
                headers=headers,
                json=chat_data
            )
            
            if response.status_code in [200, 201]:
                chat_info = response.json()
                chat_id = chat_info.get("id")
                logger.info(f"Created new chat with ID: {chat_id}")
                return chat_id
            elif response.status_code == 409:  # Conflict, chat might already exist
                try:
                    # Try to extract the existing chat ID from the error response
                    error_data = response.json()
                    if "error" in error_data and "message" in error_data["error"]:
                        # Some error messages contain the existing chat ID
                        error_msg = error_data["error"]["message"]
                        if "existing chat" in error_msg.lower():
                            # Extract the chat ID using string manipulation
                            import re
                            chat_id_match = re.search(r"'([^']+)'", error_msg)
                            if chat_id_match:
                                existing_chat_id = chat_id_match.group(1)
                                logger.info(f"Found existing chat with ID: {existing_chat_id}")
                                return existing_chat_id
                except Exception as e:
                    logger.error(f"Error parsing conflict response: {e}")
                
                # If we couldn't extract the ID, try to search for existing chats
                return self._find_existing_chat_with_user(user_id)
            else:
                logger.error(f"Failed to create chat. Status: {response.status_code}, Response: {response.text}")
                return None
        except Exception as e:
            logger.error(f"Error creating chat: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def _find_existing_chat_with_user(self, user_id):
        """Find an existing chat with the specified user"""
        if not self.access_token:
            if not self.authenticate():
                return None
                
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            # Get all chats using application permission Chat.Read.All
            response = requests.get(
                "https://graph.microsoft.com/v1.0/chats",
                headers=headers
            )
            
            if response.status_code == 200:
                chats_data = response.json()
                if "value" in chats_data:
                    # For each chat, check if the target user is a member
                    for chat in chats_data["value"]:
                        chat_id = chat.get("id")
                        
                        # Get members of this chat
                        members_response = requests.get(
                            f"https://graph.microsoft.com/v1.0/chats/{chat_id}/members",
                            headers=headers
                        )
                        
                        if members_response.status_code == 200:
                            members_data = members_response.json()
                            if "value" in members_data:
                                # Check if our target user is a member
                                for member in members_data["value"]:
                                    if member.get("userId") == user_id:
                                        logger.info(f"Found existing chat with user {user_id}, chat ID: {chat_id}")
                                        return chat_id
            
            logger.warning(f"No existing chat found with user {user_id}")
            return None
        except Exception as e:
            logger.error(f"Error finding existing chat: {e}")
            return None
    
    def create_direct_chat_alternative(self, user_id):
        """
        Alternative method to create a chat using application permissions
        """
        if not self.access_token:
            if not self.authenticate():
                return None
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Using application permissions to create a group chat
        chat_data = {
            "chatType": "group",
            "topic": "Timesheet Notification",
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_id}"
                }
            ]
        }
        
        try:
            response = requests.post(
                "https://graph.microsoft.com/v1.0/chats",
                headers=headers,
                json=chat_data
            )
            
            if response.status_code in [200, 201]:
                chat_info = response.json()
                chat_id = chat_info.get("id")
                logger.info(f"Created chat with alternative method: {chat_id}")
                return chat_id
            else:
                logger.error(f"Alternative chat creation failed. Status: {response.status_code}, Response: {response.text}")
                return None
        except Exception as e:
            logger.error(f"Error in alternative chat creation: {e}")
            return None
    
    def send_direct_message(self, chat_id, message_content):
        """Send a message to a Teams chat using application permissions"""
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
            
            response = requests.post(
                url,
                headers=headers,
                json=message_data
            )
            
            if response.status_code in [200, 201]:
                logger.info(f"Message sent successfully to chat {chat_id}")
                return True
            else:
                logger.error(f"Failed to send message. Status: {response.status_code}, Response: {response.text}")
                return False
        except Exception as e:
            logger.error(f"Error sending message: {e}")
            return False
    
    def check_user_exists(self, user_id):
        """Check if a user with the given ID exists"""
        if not self.access_token:
            if not self.authenticate():
                return None, "Authentication failed"
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            url = f"https://graph.microsoft.com/v1.0/users/{user_id}"
            logger.info(f"Checking if user exists: {user_id}")
            
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                user_data = response.json()
                logger.info(f"User found: {user_data.get('displayName', 'Unknown')}")
                return user_data, None
            else:
                logger.error(f"User not found. Status: {response.status_code}, Response: {response.text}")
                return None, f"User not found. Status: {response.status_code}"
        except Exception as e:
            error_message = str(e)
            logger.error(f"Error checking user: {error_message}")
            return None, error_message
    
    def test_organization_access(self):
        """Test if we can access organization information"""
        if not self.access_token:
            if not self.authenticate():
                return None, "Authentication failed"
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            url = "https://graph.microsoft.com/v1.0/organization"
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                return response.json(), None
            else:
                return None, f"Status: {response.status_code}, Response: {response.text}"
        except Exception as e:
            return None, str(e)
    
    def test_users_access(self):
        """Test if we can list users"""
        if not self.access_token:
            if not self.authenticate():
                return None, "Authentication failed"
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            url = "https://graph.microsoft.com/v1.0/users?$top=5"
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                return response.json(), None
            else:
                return None, f"Status: {response.status_code}, Response: {response.text}"
        except Exception as e:
            return None, str(e)
    
    def test_chats_access(self):
        """Test if we can list chats"""
        if not self.access_token:
            if not self.authenticate():
                return None, "Authentication failed"
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            url = "https://graph.microsoft.com/v1.0/chats"
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                return response.json(), None
            else:
                return None, f"Status: {response.status_code}, Response: {response.text}"
        except Exception as e:
            return None, str(e)
    
    def test_create_chat_permission(self, user_id):
        """Test chat creation with different approaches"""
        results = {}
        
        # Test application permission approach (group chat)
        group_chat_payload = {
            "chatType": "group",
            "topic": "Test Chat",
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_id}"
                }
            ]
        }
        
        results["app_permission"] = self._test_chat_creation(group_chat_payload)
        
        return results
    
    def _test_chat_creation(self, payload):
        """Helper method to test a specific chat creation payload"""
        if not self.access_token:
            if not self.authenticate():
                return {"result": None, "error": "Authentication failed"}
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            url = "https://graph.microsoft.com/v1.0/chats"
            logger.info(f"Testing chat creation with payload: {json.dumps(payload)}")
            
            response = requests.post(url, headers=headers, json=payload)
            
            if response.status_code in [200, 201]:
                result = response.json()
                logger.info(f"Chat creation test succeeded")
                return {"result": result, "error": None}
            else:
                error = f"Status: {response.status_code}, Response: {response.text}"
                logger.error(f"Chat creation test failed: {error}")
                return {"result": None, "error": error}
        except Exception as e:
            error = str(e)
            logger.error(f"Exception in chat creation test: {error}")
            return {"result": None, "error": error}