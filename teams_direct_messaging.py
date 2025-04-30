import requests
import json
import logging
import msal
import time
import traceback
import base64

# Configure extra detailed logging
logging.basicConfig(
    level=logging.DEBUG,  # Set to DEBUG for maximum detail
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

class TeamsMessenger:
    """Class focused on debugging Teams messaging issues"""
    
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
        
        # Print initialization parameters (masking sensitive data)
        logger.debug(f"Initialized TeamsMessenger with:")
        logger.debug(f"Client ID: {client_id[:5]}...{client_id[-5:] if len(client_id) > 10 else client_id}")
        logger.debug(f"Tenant ID: {tenant_id[:5]}...{tenant_id[-5:] if len(tenant_id) > 10 else tenant_id}")
    
    def authenticate(self):
        """Get access token with detailed debugging"""
        logger.debug("Authenticating with Microsoft Graph API")
        result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        
        if "access_token" in result:
            self.access_token = result["access_token"]
            token_info = self._decode_token(self.access_token)
            logger.debug(f"Successfully authenticated with token: {self.access_token[:10]}...{self.access_token[-10:] if len(self.access_token) > 20 else self.access_token}")
            logger.debug(f"Token expires: {result.get('expires_in')} seconds")
            logger.debug(f"Token scopes: {result.get('scope', 'none specified')}")
            logger.debug(f"Token app roles: {token_info.get('roles', ['none found'])}")
            return True
        else:
            logger.error(f"Authentication failed: {result.get('error')}")
            logger.error(f"Error description: {result.get('error_description')}")
            return False
    
    def _decode_token(self, token):
        """Decode JWT token to extract information"""
        try:
            # Split the token and take the payload (second part)
            parts = token.split('.')
            if len(parts) < 2:
                return {"error": "Not a valid JWT token"}
                
            # Decode payload (add padding if needed)
            payload = parts[1]
            payload += '=' * (4 - len(payload) % 4) if len(payload) % 4 != 0 else ''
            decoded_bytes = base64.b64decode(payload)
            decoded_str = decoded_bytes.decode('utf-8')
            return json.loads(decoded_str)
        except Exception as e:
            logger.error(f"Error decoding token: {e}")
            return {"error": str(e)}
    
    def notify_user(self, user_id, message_text):
        """Create a chat and debug message sending issues"""
        if not self.access_token and not self.authenticate():
            logger.error("Authentication failed")
            return False
            
        # STEP 1: Create the chat
        chat_id = self._create_chat_for_debug(user_id)
        if not chat_id:
            logger.error("Failed to create or get chat ID")
            return False
            
        # STEP 2: Debug chat details
        chat_details = self._get_chat_details(chat_id)
        
        # STEP 3: Try different message sending approaches with detailed debugging
        success = self._debug_send_message(chat_id, message_text)
        
        return success
    
    def _create_chat_for_debug(self, user_id):
        """Create a chat with detailed debugging"""
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Unique chat name for debugging
        chat_topic = f"Debug Chat {int(time.time())}"
        
        # Create chat data
        chat_data = {
            "chatType": "group",
            "topic": chat_topic,
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_id}"
                }
            ]
        }
        
        # Try both v1.0 and beta endpoints
        endpoints = [
            "https://graph.microsoft.com/v1.0/chats",
            "https://graph.microsoft.com/beta/chats"
        ]
        
        logger.debug(f"Attempting to create chat with user ID: {user_id}")
        logger.debug(f"Chat data: {json.dumps(chat_data)}")
        
        for endpoint in endpoints:
            try:
                logger.debug(f"Trying endpoint: {endpoint}")
                response = requests.post(endpoint, headers=headers, json=chat_data)
                
                logger.debug(f"Response status: {response.status_code}")
                logger.debug(f"Response headers: {dict(response.headers)}")
                
                try:
                    response_json = response.json()
                    logger.debug(f"Response JSON: {json.dumps(response_json)}")
                except:
                    logger.debug(f"Response text: {response.text}")
                
                if response.status_code in [200, 201]:
                    # Success
                    chat_id = response.json().get("id")
                    logger.debug(f"Successfully created chat with ID: {chat_id}")
                    return chat_id
                elif response.status_code == 409:  # Conflict - chat already exists
                    try:
                        # Try to extract existing chat ID
                        error_data = response.json()
                        error_message = error_data.get("error", {}).get("message", "")
                        
                        import re
                        match = re.search(r"'([^']+)'", error_message)
                        if match:
                            chat_id = match.group(1)
                            logger.debug(f"Found existing chat with ID: {chat_id}")
                            return chat_id
                    except Exception as e:
                        logger.error(f"Error parsing conflict response: {str(e)}")
                else:
                    logger.error(f"Failed with status code: {response.status_code}")
            except Exception as e:
                logger.error(f"Exception during chat creation: {str(e)}")
        
        # Both endpoints failed
        logger.error("Failed to create chat with either endpoint")
        return None
    
    def _get_chat_details(self, chat_id):
        """Get detailed information about the chat"""
        if not chat_id:
            return None
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        logger.debug(f"Getting details for chat ID: {chat_id}")
        
        try:
            # Try both endpoints
            for endpoint in ["https://graph.microsoft.com/v1.0", "https://graph.microsoft.com/beta"]:
                url = f"{endpoint}/chats/{chat_id}"
                logger.debug(f"Fetching chat details from: {url}")
                
                response = requests.get(url, headers=headers)
                
                logger.debug(f"Response status: {response.status_code}")
                
                if response.status_code == 200:
                    chat_details = response.json()
                    logger.debug(f"Chat details: {json.dumps(chat_details)}")
                    
                    # Also try to get members
                    members_url = f"{endpoint}/chats/{chat_id}/members"
                    members_response = requests.get(members_url, headers=headers)
                    
                    if members_response.status_code == 200:
                        members_data = members_response.json()
                        logger.debug(f"Chat members: {json.dumps(members_data)}")
                    else:
                        logger.debug(f"Failed to get chat members: {members_response.status_code}")
                    
                    # Also check installed apps
                    apps_url = f"{endpoint}/chats/{chat_id}/installedApps"
                    apps_response = requests.get(apps_url, headers=headers)
                    
                    if apps_response.status_code == 200:
                        apps_data = apps_response.json()
                        logger.debug(f"Chat installed apps: {json.dumps(apps_data)}")
                    else:
                        logger.debug(f"Failed to get installed apps: {apps_response.status_code}")
                        
                    return chat_details
                else:
                    logger.debug(f"Failed to get chat details with {endpoint}: {response.status_code}")
        except Exception as e:
            logger.error(f"Error getting chat details: {str(e)}")
        
        return None
    
    def _debug_send_message(self, chat_id, message_text):
        """Debug message sending with multiple attempts and formats"""
        if not chat_id:
            logger.error("Cannot send message - no chat ID provided")
            return False
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Wait for chat to initialize
        logger.debug(f"Waiting 10 seconds for chat {chat_id} to initialize...")
        time.sleep(10)
        
        # Define different message formats to try
        message_formats = [
            # Plain text format
            {
                "body": {
                    "contentType": "text",
                    "content": message_text
                }
            },
            # HTML format
            {
                "body": {
                    "contentType": "html",
                    "content": f"<p>{message_text}</p>"
                }
            },
            # Minimal format
            {
                "body": {
                    "content": "Test message"
                }
            },
            # Alternative format
            {
                "content": {
                    "body": {
                        "contentType": "text",
                        "content": "Alternative format test"
                    }
                }
            }
        ]
        
        # Try both endpoints
        endpoints = [
            "https://graph.microsoft.com/v1.0/chats/{chat_id}/messages",
            "https://graph.microsoft.com/beta/chats/{chat_id}/messages"
        ]
        
        # Try each combination of endpoint and message format
        for endpoint_template in endpoints:
            endpoint = endpoint_template.format(chat_id=chat_id)
            logger.debug(f"Trying endpoint: {endpoint}")
            
            for i, msg_format in enumerate(message_formats):
                try:
                    logger.debug(f"Trying message format {i+1}: {json.dumps(msg_format)}")
                    
                    # Wait between attempts
                    if i > 0:
                        time.sleep(3)
                    
                    response = requests.post(endpoint, headers=headers, json=msg_format)
                    
                    logger.debug(f"Response status: {response.status_code}")
                    logger.debug(f"Response headers: {dict(response.headers)}")
                    
                    try:
                        response_json = response.json()
                        logger.debug(f"Response JSON: {json.dumps(response_json)}")
                    except:
                        logger.debug(f"Response text: {response.text}")
                    
                    if response.status_code in [200, 201]:
                        logger.debug(f"Successfully sent message with format {i+1}")
                        return True
                    else:
                        logger.debug(f"Failed to send message with format {i+1}")
                        
                        # Parse error for more details
                        try:
                            error_data = response.json()
                            error_code = error_data.get("error", {}).get("code", "Unknown")
                            error_message = error_data.get("error", {}).get("message", "Unknown")
                            logger.debug(f"Error code: {error_code}")
                            logger.debug(f"Error message: {error_message}")
                        except:
                            pass
                except Exception as e:
                    logger.error(f"Exception during message send attempt: {str(e)}")
        
        # Try another approach - sending message to beta endpoint with a different payload structure
        try:
            logger.debug("Trying beta endpoint with chatMessage format")
            beta_url = f"https://graph.microsoft.com/beta/chats/{chat_id}/messages"
            
            chat_message = {
                "@odata.type": "#microsoft.graph.chatMessage",
                "body": {
                    "contentType": "text",
                    "content": "Testing message delivery (special format)"
                }
            }
            
            response = requests.post(beta_url, headers=headers, json=chat_message)
            logger.debug(f"Beta special format response: {response.status_code}")
            
            if response.status_code in [200, 201]:
                logger.debug("Successfully sent message with special beta format")
                return True
        except Exception as e:
            logger.error(f"Error with special beta format: {str(e)}")
        
        # All approaches failed
        logger.error("All message sending approaches failed")
        return False


