
import requests
import json
import msal
import time
import base64
import streamlit as st

class TeamsMessenger:
    """Simplified class with direct UI debug output"""
    
    def __init__(self, client_id, client_secret, tenant_id):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = None
        self.debug_messages = []
        
        # Create MSAL application
        self.app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=f"https://login.microsoftonline.com/{tenant_id}"
        )
        
        self.add_debug("Initialized TeamsMessenger")
    
    def add_debug(self, message, is_error=False):
        """Add a debug message to the internal list"""
        self.debug_messages.append({"message": message, "is_error": is_error})
        print(f"DEBUG: {message}")  # Also print to console
    
    def authenticate(self):
        """Get access token"""
        self.add_debug("Authenticating with Microsoft Graph API")
        result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        
        if "access_token" in result:
            self.access_token = result["access_token"]
            self.add_debug("Authentication successful!")
            
            # Decode token to check permissions
            token_info = self._decode_token(self.access_token)
            if "roles" in token_info:
                roles = token_info.get("roles", [])
                self.add_debug(f"Token roles: {', '.join(roles)}")
                
                # Check for required permissions
                required_permissions = ["Chat.Create", "Chat.ReadWrite.All", "Chat.Read.All"]
                for perm in required_permissions:
                    if any(perm in role for role in roles):
                        self.add_debug(f"✅ Found permission: {perm}")
                    else:
                        self.add_debug(f"❌ Missing permission: {perm}", True)
            else:
                self.add_debug("No roles found in token", True)
            
            return True
        else:
            self.add_debug(f"Authentication failed: {result.get('error')}", True)
            self.add_debug(f"Error description: {result.get('error_description')}", True)
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
            self.add_debug(f"Error decoding token: {e}", True)
            return {"error": str(e)}
    
    def notify_user(self, user_id, message_text):
        """Attempt to send a message with detailed debugging"""
        if not self.access_token and not self.authenticate():
            self.add_debug("Authentication failed", True)
            return False
            
        # Step 1: Create chat
        self.add_debug("----- STEP 1: Create Chat -----")
        chat_id = self._create_chat(user_id)
        if not chat_id:
            self.add_debug("Failed to get a valid chat ID", True)
            return False
        
        # Step 2: Send message with multiple attempts
        self.add_debug("----- STEP 2: Send Message -----")
        return self._try_all_message_methods(chat_id, message_text)
    
    def _create_chat(self, user_id):
        """Create a chat and return the ID"""
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Create chat data
        chat_topic = f"Chat {int(time.time())}"
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
        
        self.add_debug(f"Creating chat with topic: {chat_topic}")
        
        try:
            # Try beta endpoint first
            url = "https://graph.microsoft.com/beta/chats"
            self.add_debug(f"POST {url}")
            
            response = requests.post(url, headers=headers, json=chat_data)
            self.add_debug(f"Response status: {response.status_code}")
            
            # Handle response
            if response.status_code in [200, 201]:
                chat_id = response.json().get("id")
                self.add_debug(f"Chat created successfully with ID: {chat_id}")
                return chat_id
            elif response.status_code == 409:  # Conflict - chat already exists
                try:
                    # Try to extract existing chat ID
                    error_data = response.json()
                    error_message = error_data.get("error", {}).get("message", "")
                    self.add_debug(f"Chat conflict: {error_message}")
                    
                    # Extract chat ID from error message
                    import re
                    match = re.search(r"'([^']+)'", error_message)
                    if match:
                        chat_id = match.group(1)
                        self.add_debug(f"Found existing chat with ID: {chat_id}")
                        return chat_id
                except Exception as e:
                    self.add_debug(f"Error parsing conflict response: {str(e)}", True)
            else:
                self.add_debug(f"Failed to create chat: {response.status_code}", True)
                try:
                    error_data = response.json()
                    error_message = error_data.get("error", {}).get("message", "")
                    self.add_debug(f"Error message: {error_message}", True)
                except:
                    self.add_debug(f"Response: {response.text[:500]}", True)
                    
        except Exception as e:
            self.add_debug(f"Exception creating chat: {str(e)}", True)
        
        # If we get here, try v1.0 endpoint
        try:
            url = "https://graph.microsoft.com/v1.0/chats"
            self.add_debug(f"Falling back to v1.0 endpoint: {url}")
            
            response = requests.post(url, headers=headers, json=chat_data)
            self.add_debug(f"v1.0 response status: {response.status_code}")
            
            if response.status_code in [200, 201]:
                chat_id = response.json().get("id")
                self.add_debug(f"Chat created successfully with v1.0 endpoint: {chat_id}")
                return chat_id
            else:
                self.add_debug(f"Failed to create chat with v1.0 endpoint: {response.status_code}", True)
        except Exception as e:
            self.add_debug(f"Exception with v1.0 endpoint: {str(e)}", True)
        
        return None
    
    def _try_all_message_methods(self, chat_id, message_text):
        """Try all possible message sending methods"""
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Wait for chat to initialize
        wait_seconds = 10
        self.add_debug(f"Waiting {wait_seconds} seconds for chat to initialize...")
        time.sleep(wait_seconds)
        
        # Try each method in sequence
        methods = [
            self._try_beta_text_message,
            self._try_v1_text_message,
            self._try_beta_html_message,
            self._try_beta_simple_message,
            self._try_welcome_message_approach
        ]
        
        for i, method in enumerate(methods):
            self.add_debug(f"Attempt {i+1}: Trying {method.__name__}")
            success = method(chat_id, message_text, headers)
            
            if success:
                self.add_debug(f"✅ {method.__name__} succeeded!")
                return True
            
            self.add_debug(f"❌ {method.__name__} failed")
            # Wait between attempts
            if i < len(methods) - 1:
                time.sleep(3)
        
        # All methods failed
        self.add_debug("All message sending methods failed", True)
        return False
    
    def _try_beta_text_message(self, chat_id, message_text, headers):
        """Try sending a message with beta endpoint, text format"""
        url = f"https://graph.microsoft.com/beta/chats/{chat_id}/messages"
        data = {
            "body": {
                "contentType": "text",
                "content": message_text
            }
        }
        
        try:
            self.add_debug(f"POST {url} (Beta Text)")
            response = requests.post(url, headers=headers, json=data)
            self.add_debug(f"Response status: {response.status_code}")
            
            if response.status_code in [200, 201]:
                return True
            else:
                try:
                    error_data = response.json()
                    error_code = error_data.get("error", {}).get("code", "Unknown")
                    error_message = error_data.get("error", {}).get("message", "Unknown") 
                    self.add_debug(f"Error code: {error_code}", True)
                    self.add_debug(f"Error message: {error_message}", True)
                except:
                    self.add_debug(f"Response: {response.text[:300]}", True)
                return False
        except Exception as e:
            self.add_debug(f"Exception: {str(e)}", True)
            return False
    
    def _try_v1_text_message(self, chat_id, message_text, headers):
        """Try sending a message with v1.0 endpoint, text format"""
        url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
        data = {
            "body": {
                "contentType": "text",
                "content": message_text
            }
        }
        
        try:
            self.add_debug(f"POST {url} (v1.0 Text)")
            response = requests.post(url, headers=headers, json=data)
            self.add_debug(f"Response status: {response.status_code}")
            
            if response.status_code in [200, 201]:
                return True
            else:
                try:
                    error_data = response.json()
                    error_code = error_data.get("error", {}).get("code", "Unknown")
                    error_message = error_data.get("error", {}).get("message", "Unknown") 
                    self.add_debug(f"Error code: {error_code}", True)
                    self.add_debug(f"Error message: {error_message}", True)
                except:
                    self.add_debug(f"Response: {response.text[:300]}", True)
                return False
        except Exception as e:
            self.add_debug(f"Exception: {str(e)}", True)
            return False
    
    def _try_beta_html_message(self, chat_id, message_text, headers):
        """Try sending a message with beta endpoint, HTML format"""
        url = f"https://graph.microsoft.com/beta/chats/{chat_id}/messages"
        data = {
            "body": {
                "contentType": "html",
                "content": f"<p>{message_text}</p>"
            }
        }
        
        try:
            self.add_debug(f"POST {url} (Beta HTML)")
            response = requests.post(url, headers=headers, json=data)
            self.add_debug(f"Response status: {response.status_code}")
            
            if response.status_code in [200, 201]:
                return True
            else:
                try:
                    error_data = response.json()
                    error_code = error_data.get("error", {}).get("code", "Unknown")
                    error_message = error_data.get("error", {}).get("message", "Unknown") 
                    self.add_debug(f"Error code: {error_code}", True)
                    self.add_debug(f"Error message: {error_message}", True)
                except:
                    self.add_debug(f"Response: {response.text[:300]}", True)
                return False
        except Exception as e:
            self.add_debug(f"Exception: {str(e)}", True)
            return False
    
    def _try_beta_simple_message(self, chat_id, message_text, headers):
        """Try sending a simple message with beta endpoint"""
        url = f"https://graph.microsoft.com/beta/chats/{chat_id}/messages"
        data = {
            "body": {
                "content": "Simple test message"
            }
        }
        
        try:
            self.add_debug(f"POST {url} (Beta Simple)")
            response = requests.post(url, headers=headers, json=data)
            self.add_debug(f"Response status: {response.status_code}")
            
            if response.status_code in [200, 201]:
                return True
            else:
                try:
                    error_data = response.json()
                    error_code = error_data.get("error", {}).get("code", "Unknown")
                    error_message = error_data.get("error", {}).get("message", "Unknown") 
                    self.add_debug(f"Error code: {error_code}", True)
                    self.add_debug(f"Error message: {error_message}", True)
                except:
                    self.add_debug(f"Response: {response.text[:300]}", True)
                return False
        except Exception as e:
            self.add_debug(f"Exception: {str(e)}", True)
            return False
    
    def _try_welcome_message_approach(self, chat_id, message_text, headers):
        """Try creating a new chat with welcome message"""
        url = "https://graph.microsoft.com/beta/chats"
        
        # Extract user ID from existing chat
        self.add_debug("Trying to get members of existing chat to create welcome message")
        
        try:
            # Get members of existing chat
            members_url = f"https://graph.microsoft.com/beta/chats/{chat_id}/members"
            members_response = requests.get(members_url, headers=headers)
            
            if members_response.status_code == 200:
                members_data = members_response.json()
                
                # Find the user that's not the app
                user_id = None
                for member in members_data.get("value", []):
                    if member.get("userId"):
                        user_id = member.get("userId")
                        self.add_debug(f"Found user ID: {user_id}")
                        break
                
                if user_id:
                    # Create a new chat with welcome message
                    chat_data = {
                        "chatType": "group",
                        "topic": f"Alert {int(time.time())}",
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
                    
                    self.add_debug(f"Creating new chat with welcome message")
                    response = requests.post(url, headers=headers, json=chat_data)
                    
                    self.add_debug(f"Response status: {response.status_code}")
                    
                    if response.status_code in [200, 201]:
                        return True
                    else:
                        self.add_debug(f"Failed to create chat with welcome message", True)
                        return False
                else:
                    self.add_debug("Could not find user ID in chat members", True)
                    return False
            else:
                self.add_debug(f"Failed to get chat members: {members_response.status_code}", True)
                return False
        except Exception as e:
            self.add_debug(f"Exception in welcome message approach: {str(e)}", True)
            return False
