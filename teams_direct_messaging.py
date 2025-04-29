import requests
import os
import streamlit as st
import json
import logging
import msal  # You'll need to pip install msal
import traceback

# Configure logging
logger = logging.getLogger(__name__)

SECOND_MEMBER_ID = (
    os.getenv("TEAMS_SECOND_MEMBER_ID") or
    st.secrets.get("TEAMS_SECOND_MEMBER", {}).get("ID")
)
if not SECOND_MEMBER_ID:
    raise RuntimeError(
        "SECOND_MEMBER_ID not found. "
        "Add it to Streamlit secrets or set the env-var TEAMS_SECOND_MEMBER_ID."
    )


class TeamsDirectMessaging:
    """Class to handle Microsoft Teams direct messaging via Graph API"""
    
    def __init__(self, client_id, client_secret, tenant_id, access_token=None):
        """Initialize with Azure AD credentials"""
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = access_token
        
        # Log initialization (mask sensitive parts)
        client_id_masked = client_id[:5] + "..." + client_id[-5:] if client_id else "None"
        tenant_id_masked = tenant_id[:5] + "..." + tenant_id[-5:] if tenant_id else "None"
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
        For delegated mode we’ve already injected self.access_token
        from app.py → nothing to do. We keep the method so the
        rest of the code can still call .authenticate().
        """
        return bool(self.access_token)
    
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
                user_id = user_data.get("id")
                logger.info(f"Found user ID: {user_id[:5]}...{user_id[-5:]}" if user_id else "No user ID found")
                return user_id
            else:
                logger.error(f"Failed to get user ID for {email}. Status: {response.status_code}, Response: {response.text}")
                return None
        except Exception as e:
            logger.error(f"Error getting user ID: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def create_chat(self, primary_user_id):
        """
        Create (or reuse) a 2-member group chat
        • primary_user_id  = the designer being notified
        • SECOND_MEMBER_ID = you (Ibrahim) for audit
        Returns chat_id or None.
        """
        # ensure we have a token
        if not self.access_token and not self.authenticate():
            return None

        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type":  "application/json"
        }

        chat_data = {
            "chatType": "group",
            "members": [
                {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["member"],
                "user@odata.bind":
                f"https://graph.microsoft.com/v1.0/users('{primary_user_id}')"
                },
                {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["member"],
                "user@odata.bind":
                f"https://graph.microsoft.com/v1.0/users('{SECOND_MEMBER_ID}')"
                }
            ]
        }

        resp = requests.post("https://graph.microsoft.com/v1.0/chats",
                            headers=headers, json=chat_data, timeout=10)

        if resp.status_code in (200, 201):
            chat_id = resp.json().get("id")
            logger.info("Group chat created: %s", chat_id)
            return chat_id

        if resp.status_code == 409:        # roster already exists
            try:
                return resp.json().get("chatId") or resp.json()["value"][0]["id"]
            except Exception:
                logger.warning("409 returned but could not parse chatId")
                return None

        logger.error("create_chat failed %s: %s", resp.status_code, resp.text)
        return None
        
    def create_direct_chat_alternative(self, user_id):
        """Try an alternative chat creation method"""
        if not self.access_token:
            if not self.authenticate():
                logger.error("Authentication failed before create_direct_chat_alternative")
                return None
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Try creating a direct message with simpler payload
        chat_data = {
            "chatType": "oneOnOne",
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["member"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_id}"
                }
            ]
        }
        
        try:
            url = "https://graph.microsoft.com/v1.0/chats"
            logger.info(f"Making POST request to alternative endpoint: {url}")
            logger.info(f"Alternative request payload: {json.dumps(chat_data)}")
            
            response = requests.post(
                url,
                headers=headers,
                json=chat_data
            )
            
            # Log status code and partial response
            logger.info(f"Alternative response status: {response.status_code}")
            if response.status_code != 201 and response.status_code != 200:
                logger.error(f"Alternative response text: {response.text[:200]}...")
            
            if response.status_code in [200, 201]:
                chat_info = response.json()
                chat_id = chat_info.get("id")
                logger.info(f"Created/found chat with alternative method ID: {chat_id}")
                return chat_id
            else:
                logger.error(f"Alternative method failed to create chat. Status: {response.status_code}")
                return None
        except Exception as e:
            logger.error(f"Error in alternative chat creation: {e}")
            logger.error(traceback.format_exc())
            return None
    def send_direct_message(self, chat_id, message_content):
        """Send a message to a Teams chat"""
        if not self.access_token:
            logger.info("No access token found, authenticating first")
            if not self.authenticate():
                logger.error("Authentication failed before send_direct_message")
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
            logger.info(f"Making POST request to: {url}")
            logger.info(f"Message length: {len(message_content)} characters")
            
            response = requests.post(
                url,
                headers=headers,
                json=message_data
            )
            
            # Log status code and partial response
            logger.info(f"Response status: {response.status_code}")
            if response.status_code != 201 and response.status_code != 200:
                logger.error(f"Response text: {response.text[:200]}...")
            
            if response.status_code in [200, 201]:
                logger.info(f"Message sent successfully to chat {chat_id}")
                return True
            else:
                logger.error(f"Failed to send message. Status: {response.status_code}, Response: {response.text}")
                return False
        except Exception as e:
            logger.error(f"Error sending message: {e}")
            logger.error(traceback.format_exc())
            return False
        
    def debug_api_call(self, method, url, payload=None, max_content_length=500):
        """Make an API call to Microsoft Graph with detailed debugging"""
        if not self.access_token:
            if not self.authenticate():
                return None, "Authentication failed"
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            logger.info(f"DEBUG: Making {method} request to: {url}")
            if payload:
                logger.info(f"DEBUG: Request payload: {json.dumps(payload)}")
                
            if method.upper() == "GET":
                response = requests.get(url, headers=headers)
            elif method.upper() == "POST":
                response = requests.post(url, headers=headers, json=payload)
            else:
                return None, f"Unsupported method: {method}"
            
            logger.info(f"DEBUG: Response status: {response.status_code}")
            logger.info(f"DEBUG: Response headers: {dict(response.headers)}")
            
            content = response.text[:max_content_length] + "..." if len(response.text) > max_content_length else response.text
            logger.info(f"DEBUG: Response content: {content}")
            
            if response.status_code >= 400:
                error_detail = f"Error {response.status_code}: {content}"
                return None, error_detail
            
            try:
                return response.json(), None
            except:
                return response.text, None
        except Exception as e:
            error_detail = f"Request failed: {str(e)}"
            logger.error(error_detail)
            logger.error(traceback.format_exc())
            return None, error_detail
        
    def test_organization_access(self):
        """Test if we can access organization information"""
        result, error = self.debug_api_call("GET", "https://graph.microsoft.com/v1.0/organization")
        return result, error

    def test_users_access(self):
        """Test if we can list users"""
        result, error = self.debug_api_call("GET", "https://graph.microsoft.com/v1.0/users?$top=5")
        return result, error

    def test_chats_access(self):
        """Test if we can list chats"""
        result, error = self.debug_api_call("GET", "https://graph.microsoft.com/v1.0/chats")
        return result, error

    def test_create_chat_permission(self, user_id):
        """Test various chat creation approaches"""
        # Try standard format
        payload1 = {
            "chatType": "oneOnOne",
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["member"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_id}')"
                }
            ]
        }
        
        result1, error1 = self.debug_api_call("POST", "https://graph.microsoft.com/v1.0/chats", payload1)
        
        # Try alternative format
        payload2 = {
            "chatType": "oneOnOne",
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["member"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_id}"
                }
            ]
        }
        
        result2, error2 = self.debug_api_call("POST", "https://graph.microsoft.com/v1.0/chats", payload2)
        
        return {
            "standard": {"result": result1, "error": error1},
            "alternative": {"result": result2, "error": error2}
        }

    def check_user_exists(self, user_id):
        """Check if a user with the given ID exists"""
        result, error = self.debug_api_call("GET", f"https://graph.microsoft.com/v1.0/users/{user_id}")
        return result, error