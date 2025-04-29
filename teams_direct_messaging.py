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
    
    def __init__(self, client_id, client_secret, tenant_id):
        """Initialize with Azure AD credentials"""
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = None
        
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
        """Authenticate with Microsoft Graph API using client credentials flow"""
        try:
            if not self.app:
                logger.error("Cannot authenticate - MSAL application not initialized")
                return False
                
            logger.info("Requesting token with scope: https://graph.microsoft.com/.default")
            
            # Acquire token for application
            result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                token_preview = self.access_token[:10] + "..." + self.access_token[-10:] if self.access_token else "None"
                logger.info(f"Successfully acquired token: {token_preview}")
                return True
            else:
                error = result.get('error', 'Unknown error')
                description = result.get('error_description', 'No description')
                logger.error(f"Authentication error: {error}: {description}")
                return False
        except Exception as e:
            logger.error(f"Error during authentication: {e}")
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
            url = f"https://graph.microsoft.com/v1.0/users/{email}"
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
                "roles": ["owner"],
                "user@odata.bind":
                f"https://graph.microsoft.com/v1.0/users('{primary_user_id}')"
                },
                {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
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
                    "roles": ["owner"],
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