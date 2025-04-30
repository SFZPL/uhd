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
        Create a group chat with a user ID 
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
        
        # Create a unique chat topic with timestamp to avoid conflicts
        chat_topic = f"Missing Timesheet Notification {int(time.time())}"
        
        # For application permissions, we need to create a group chat
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
        
        try:
            url = "https://graph.microsoft.com/v1.0/chats"
            logger.info(f"Creating chat with user ID: {user_id} and topic: {chat_topic}")
            
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
                logger.warning("Chat already exists, trying to extract ID from error message")
                try:
                    # Try to get the existing chat ID from the response
                    error_data = response.json()
                    error_message = error_data.get("error", {}).get("message", "")
                    
                    # Try to extract the chat ID from the error message
                    if "'" in error_message:
                        import re
                        match = re.search(r"'([^']+)'", error_message)
                        if match:
                            chat_id = match.group(1)
                            logger.info(f"Extracted existing chat ID from error: {chat_id}")
                            return chat_id
                except Exception as e:
                    logger.error(f"Error parsing conflict response: {e}")
                
                # If we couldn't extract the ID, try the alternative method
                logger.info("Could not extract chat ID from error, trying alternative method")
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
        Alternative approach to create a chat with a different format
        """
        if not self.access_token:
            if not self.authenticate():
                return None
                
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Create a unique chat topic with timestamp to avoid conflicts
        chat_topic = f"Timesheet Alert {int(time.time())}"
        
        # Alternative format for chat creation
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
        
        try:
            # Try with beta endpoint
            url = "https://graph.microsoft.com/beta/chats"
            logger.info(f"Trying alternative chat creation with beta endpoint for user ID: {user_id}")
            
            response = requests.post(url, headers=headers, json=chat_data)
            
            if response.status_code in [200, 201]:
                data = response.json()
                chat_id = data.get("id")
                logger.info(f"Successfully created chat with beta endpoint: {chat_id}")
                return chat_id
            else:
                logger.error(f"Beta endpoint chat creation failed: {response.status_code} {response.text}")
                
                # Try one more approach - chatType oneOnOne with beta endpoint
                try:
                    one_on_one_data = {
                        "chatType": "oneOnOne",
                        "members": [
                            {
                                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                                "roles": ["owner"],
                                "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_id}"
                            }
                        ]
                    }
                    
                    final_response = requests.post(url, headers=headers, json=one_on_one_data)
                    
                    if final_response.status_code in [200, 201]:
                        final_data = final_response.json()
                        chat_id = final_data.get("id")
                        logger.info(f"Successfully created oneOnOne chat with beta endpoint: {chat_id}")
                        return chat_id
                    else:
                        logger.error(f"Final attempt failed: {final_response.status_code} {final_response.text}")
                        return None
                except Exception as e:
                    logger.error(f"Error in final chat creation attempt: {e}")
                    return None
        except Exception as e:
            logger.error(f"Error in _create_chat_alternative: {e}")
            logger.error(traceback.format_exc())
            return None
    
    def send_direct_message(self, chat_id, message_content):
        """
        Send a message to a chat using the beta endpoint
        """
        if not self.access_token:
            if not self.authenticate():
                return False
                
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Ensure we have chat_id
        if not chat_id:
            logger.error("Cannot send message - chat_id is empty")
            return False
            
        # Wait to ensure chat is fully initialized
        logger.info(f"Waiting 5 seconds for chat {chat_id} to fully initialize...")
        time.sleep(5)
        
        # Use the beta endpoint
        url = f"https://graph.microsoft.com/beta/chats/{chat_id}/messages"
        
        # Prepare a simple text message
        message_data = {
            "body": {
                "content": message_content,
                "contentType": "text"
            }
        }
        
        try:
            logger.info(f"Sending message to beta endpoint for chat ID: {chat_id}")
            logger.info(f"Message content preview: {message_content[:100]}...")
            
            response = requests.post(url, headers=headers, json=message_data)
            
            logger.info(f"Send message response status: {response.status_code}")
            logger.info(f"Response headers: {dict(response.headers)}")
            logger.info(f"Response content (truncated): {response.text[:300]}...")
            
            if response.status_code in [200, 201]:
                logger.info(f"Successfully sent message to chat {chat_id}")
                return True
            else:
                logger.error(f"Failed to send message: {response.status_code} {response.text}")
                
                # As a last resort, try adding a team member to the chat
                # Sometimes this "wakes up" the chat functionality
                try:
                    logger.info("Attempting to add a bot to the chat to wake it up")
                    # This is a common Microsoft bot ID - adjust if needed
                    bot_id = "28:0d5e1277-895a-4a41-bca4-6e5e2e0f8e36"
                    
                    add_member_url = f"https://graph.microsoft.com/beta/chats/{chat_id}/members"
                    add_member_data = {
                        "@odata.type": "#microsoft.graph.aadUserConversationMember",
                        "roles": ["member"],
                        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{bot_id}')"
                    }
                    
                    requests.post(add_member_url, headers=headers, json=add_member_data)
                    
                    # Try sending the message again
                    time.sleep(2)
                    retry_response = requests.post(url, headers=headers, json=message_data)
                    
                    if retry_response.status_code in [200, 201]:
                        logger.info("Successfully sent message after adding bot")
                        return True
                except Exception as e:
                    logger.error(f"Error in bot addition attempt: {e}")
                
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


def send_teams_direct_message(
        designer_name: str,
        designer_teams_id: str,
        tasks: list,
        teams_client: TeamsDirectMessaging
    ):
    """
    Send a direct message to a designer in Teams about missing timesheet entries
    """
    try:
        # Add detailed debug logs
        logger.info(f"Attempting to send message to {designer_name} with ID {designer_teams_id}")
        
        # Test authentication first
        logger.info("Authenticating with Microsoft Graph API...")
        if not teams_client.authenticate():
            logger.error("Authentication failed with Microsoft Graph API")
            return False
        
        logger.info("Authentication successful, token acquired")
        
        # Try to create a chat
        logger.info(f"Attempting to create chat with user ID: {designer_teams_id}")
        chat_id = teams_client.create_chat(designer_teams_id)
            
        if not chat_id:
            logger.error(f"Failed to create chat with user {designer_name}")
            return False
            
        logger.info(f"Successfully created/found chat with ID: {chat_id}")
        
        # Format a simple text message - avoid HTML completely
        max_days_overdue = max(t.get("Days Overdue", 0) for t in tasks)
        one_day = (max_days_overdue == 1)
        
        message = f"{'🟠' if one_day else '🔴'} Missing Timesheet Alert\n\n"
        message += f"Hi {designer_name},\n\n"
        
        if one_day:
            message += "This is a gentle reminder to log your hours for the task(s) below — it only takes a minute:\n\n"
        else:
            message += "It looks like no hours have been logged for the past two days for the task(s) below:\n\n"
        
        # Add tasks in simple text format
        for i, t in enumerate(tasks, 1):
            message += f"{i}. Project: {t.get('Project', 'Unknown')}\n"
            message += f"   Task: {t.get('Task', 'Unknown')}\n"
            message += f"   Date: {t.get('Date', '—')}\n"
            message += f"   CS Contact: {t.get('Client Success Member', 'Unknown')}\n\n"
        
        # Add footer
        if one_day:
            message += "Taking a minute now helps us stay on top of things later 🙌\n"
            message += "Let us know if you need any support with this.\n\n"
        else:
            message += "We completely understand things can get busy — but consistent time logging "
            message += "helps us improve project planning and smooth reporting.\n"
            message += "If something's holding you back from logging your hours, just reach out. We're here to help.\n\n"
        
        message += "— Automated notice from the Missing Timesheet Reporter"
        
        # Send the message
        message_sent = teams_client.send_direct_message(chat_id, message)
        
        if message_sent:
            logger.info(f"Direct message sent to {designer_name} via Teams")
            return True
        else:
            logger.error(f"Failed to send direct message to {designer_name}")
            return False

    except Exception as exc:
        logger.error(f"send_teams_direct_message failed: {exc}", exc_info=True)
        # Log the full exception traceback for debugging
        logger.error(traceback.format_exc())
        return False


def render_teams_direct_messaging_ui():
    """Render the UI for Teams direct messaging configuration"""
    with st.sidebar.expander("Teams Direct Messaging", expanded=False):
        st.session_state.teams_direct_msg_enabled = st.checkbox(
            "Enable Teams Direct Messages", 
            value=st.session_state.teams_direct_msg_enabled,
            help="Send personal chat messages to designers in Microsoft Teams"
        )
        
        # Azure AD App registration details
        st.markdown("### Azure AD App Configuration")
        st.info("You need to register an app in Azure AD with Microsoft Graph API permissions: Chat.Create, Chat.Read.All, Chat.ReadWrite.All.")
        
        st.session_state.azure_client_id = st.text_input(
            "Azure AD Client ID",
            value=st.session_state.azure_client_id,
            type="password", 
            help="Client ID from your Azure AD app registration"
        )
        
        st.session_state.azure_client_secret = st.text_input(
            "Azure AD Client Secret",
            value=st.session_state.azure_client_secret,
            type="password",
            help="Client secret from your Azure AD app registration"
        )
        
        st.session_state.azure_tenant_id = st.text_input(
            "Azure AD Tenant ID",
            value=st.session_state.azure_tenant_id,
            type="password",
            help="Tenant ID of your Azure AD"
        )
        
        # Simple authentication test
        if st.button("Test Authentication"):
            if not (st.session_state.azure_client_id and st.session_state.azure_client_secret and st.session_state.azure_tenant_id):
                st.error("Please configure Azure AD credentials first")
            else:
                try:
                    # Create Teams client
                    client = TeamsDirectMessaging(
                        st.session_state.azure_client_id,
                        st.session_state.azure_client_secret,
                        st.session_state.azure_tenant_id
                    )
                    
                    # Test authentication
                    with st.spinner("Testing authentication..."):
                        auth_result = client.authenticate()
                        if auth_result:
                            st.success("✅ Authentication successful!")
                            
                            # Try to get token info
                            token_info = client.check_authentication()
                            if token_info["status"] == "success":
                                st.subheader("Token Information")
                                st.write(f"App ID: {token_info.get('app_id')}")
                                st.write(f"Tenant ID: {token_info.get('tenant_id')}")
                                
                                # Show permissions
                                if 'roles' in token_info and token_info['roles']:
                                    st.write("Application Permissions:")
                                    for role in token_info['roles']:
                                        st.write(f"- {role}")
                        else:
                            st.error("❌ Authentication failed!")
                except Exception as e:
                    st.error(f"Error testing authentication: {str(e)}")
        
        # Designer to Teams ID mapping
        st.markdown("### Designer Teams User ID Mapping")
        st.markdown("""
        Map designer names to their Microsoft Teams user IDs.
        
        You can get user IDs from:
        1. Microsoft Teams admin portal
        2. Using the Microsoft Graph Explorer
        3. Using email addresses (if your app has User.Read.All permission)
        """)
        
        # Allow mapping via email
        st.markdown("#### Map Designer by Email Address")
        col1, col2 = st.columns(2)
        with col1:
            lookup_designer = st.text_input("Designer Name", key="lookup_designer_name")
        with col2:
            lookup_email = st.text_input("Designer Email", key="lookup_designer_email")
        
        if st.button("Look up Teams ID"):
            if not (st.session_state.azure_client_id and st.session_state.azure_client_secret and st.session_state.azure_tenant_id):
                st.error("Please configure Azure AD credentials first")
            elif not (lookup_designer and lookup_email):
                st.error("Please enter both designer name and email address")
            else:
                # Create Teams client for lookup
                client = TeamsDirectMessaging(
                    st.session_state.azure_client_id,
                    st.session_state.azure_client_secret,
                    st.session_state.azure_tenant_id
                )
                
                with st.spinner("Looking up Teams user ID..."):
                    # Authenticate
                    if not client.authenticate():
                        st.error("Failed to authenticate with Microsoft Graph API")
                    else:
                        # Look up user ID by email
                        user_id = client.get_user_id_by_email(lookup_email)
                        
                        if user_id:
                            st.session_state.designer_teams_id_mapping[lookup_designer] = user_id
                            st.success(f"Found Teams ID for {lookup_designer}!")
                        else:
                            st.error(f"Could not find Teams user with email {lookup_email}. This might be because your app doesn't have User.Read.All permission.")
                            st.info("You'll need to manually add the designer's Teams ID.")
        
        # Manual mapping
        st.markdown("#### Manual Mapping")
        col1, col2 = st.columns(2)
        with col1:
            new_designer = st.text_input("Designer Name", key="new_teams_designer")
        with col2:
            new_teams_id = st.text_input("Teams User ID", key="new_teams_user_id")
        
        if st.button("Add Mapping"):
            if new_designer and new_teams_id:
                st.session_state.designer_teams_id_mapping[new_designer] = new_teams_id
                st.success(f"Added Teams ID mapping for {new_designer}")
            else:
                st.error("Please enter both designer name and Teams user ID")
        
        # Display current mappings and allow removal
        if st.session_state.designer_teams_id_mapping:
            st.markdown("### Current Mappings")
            for idx, (designer, teams_id) in enumerate(st.session_state.designer_teams_id_mapping.items()):
                col1, col2, col3 = st.columns([3, 3, 1])
                with col1:
                    st.text(designer)
                with col2:
                    # Show just part of the ID for security
                    masked_id = teams_id[:5] + "..." + teams_id[-5:] if len(teams_id) > 10 else teams_id
                    st.text(masked_id)
                with col3:
                    if st.button("Remove", key=f"remove_teams_{idx}"):
                        del st.session_state.designer_teams_id_mapping[designer]
                        st.experimental_rerun()
        
        # Test message section
        st.markdown("### Test Direct Message")
        test_designer = st.selectbox(
            "Select Designer to Test", 
            options=list(st.session_state.designer_teams_id_mapping.keys()) if st.session_state.designer_teams_id_mapping else ["No designers mapped"],
            key="teams_direct_msg_test_designer"
        )
        
        if st.button("Send Test Message"):
            if not st.session_state.designer_teams_id_mapping:
                st.error("Please add at least one designer Teams ID mapping")
            elif not (st.session_state.azure_client_id and st.session_state.azure_client_secret and st.session_state.azure_tenant_id):
                st.error("Please configure Azure AD credentials first")
            elif test_designer == "No designers mapped":
                st.error("Please add at least one designer mapping first")
            else:
                # Get test designer Teams ID
                teams_id = st.session_state.designer_teams_id_mapping.get(test_designer)
                
                # Create Teams client for testing
                client = TeamsDirectMessaging(
                    st.session_state.azure_client_id,
                    st.session_state.azure_client_secret,
                    st.session_state.azure_tenant_id
                )
                
                # Create test task
                test_task = [{
                    "Project": "Test Project",
                    "Task": "Test Task",
                    "Start Time": "09:00",
                    "End Time": "17:00",
                    "Allocated Hours": 8.0,
                    "Date": datetime.now().date().strftime("%Y-%m-%d"),
                    "Days Overdue": 1,
                    "Client Success Member": "Test Manager"
                }]
                
                with st.spinner("Sending test message..."):
                    # Send the test message
                    message_sent = send_teams_direct_message(
                        test_designer,
                        teams_id,
                        test_task,
                        client
                    )
                    
                    if message_sent:
                        st.success(f"Test message sent to {test_designer}")
                    else:
                        st.error(f"Failed to send test message to {test_designer}")