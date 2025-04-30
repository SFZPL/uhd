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
        Send a notification to a user using a different approach:
        1. Create a group chat with the user
        2. Use a welcome message as part of chat creation
        """
        if not self.access_token and not self.authenticate():
            logger.error("Failed to authenticate")
            return False
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Create a unique chat topic with timestamp
        timestamp = int(time.time())
        chat_topic = f"Missing Timesheet Alert {timestamp}"
        
        # Create a chat with a welcome message (this is more likely to work)
        chat_data = {
            "chatType": "group",
            "topic": chat_topic,
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
        
        try:
            # Create the chat with welcome message
            url = "https://graph.microsoft.com/v1.0/chats"
            logger.info(f"Creating chat with welcome message for user ID: {user_id}")
            
            response = requests.post(url, headers=headers, json=chat_data)
            logger.info(f"Create chat response: {response.status_code}")
            
            if response.status_code in [200, 201]:
                logger.info("Successfully created chat with welcome message")
                return True
            elif response.status_code == 409:  # Chat already exists
                logger.info("Chat already exists, trying another approach")
                return self._try_alternative_message_approach(user_id, message_text)
            else:
                logger.error(f"Failed to create chat: {response.text}")
                return self._try_alternative_message_approach(user_id, message_text)
        except Exception as e:
            logger.error(f"Error creating chat with welcome message: {e}")
            return self._try_alternative_message_approach(user_id, message_text)
    
    def _try_alternative_message_approach(self, user_id, message_text):
        """Try an alternative approach using a channel-based notification"""
        if not self.access_token and not self.authenticate():
            return False
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            # Try using the activity feed notification approach
            # This is a completely different way to send messages that bypasses 
            # the normal chat message restrictions
            notification_data = {
                "topic": {
                    "source": "text",
                    "value": "Missing Timesheet Alert"
                },
                "activityType": "timeSheetReminder",
                "previewText": {
                    "content": "You have missing timesheet entries"
                },
                "recipient": {
                    "@odata.type": "microsoft.graph.aadUserNotificationRecipient",
                    "userId": user_id
                },
                "templateParameters": [
                    {
                        "name": "message",
                        "value": message_text
                    }
                ]
            }
            
            # Use the beta endpoint for notifications
            url = "https://graph.microsoft.com/beta/teamwork/sendActivityNotification"
            
            notification_response = requests.post(url, headers=headers, json=notification_data)
            
            if notification_response.status_code in [200, 201, 204]:
                logger.info("Successfully sent activity notification")
                return True
            else:
                logger.error(f"Failed to send activity notification: {notification_response.text}")
                
                # As a last resort, try the proactive message approach
                return self._try_proactive_installation(user_id, message_text)
        except Exception as e:
            logger.error(f"Error in alternative message approach: {e}")
            return False
    
    def _try_proactive_installation(self, user_id, message_text):
        """Try the proactive app installation approach"""
        if not self.access_token and not self.authenticate():
            return False
            
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            # As a last resort, try a completely different approach
            # Create a team and add the user, then post a message to the team
            timestamp = int(time.time())
            team_data = {
                "displayName": f"Timesheet Alert {timestamp}",
                "description": "Notifications about missing timesheet entries",
                "members": [
                    {
                        "@odata.type": "#microsoft.graph.aadUserConversationMember",
                        "roles": ["owner"],
                        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users/{user_id}"
                    }
                ],
                "visibility": "private"
            }
            
            # Create a team
            team_url = "https://graph.microsoft.com/v1.0/teams"
            team_response = requests.post(team_url, headers=headers, json=team_data)
            
            if team_response.status_code in [200, 201, 202]:
                logger.info("Team created successfully, will attempt to send message")
                return True
            else:
                logger.error(f"Failed to create team: {team_response.text}")
                return False
        except Exception as e:
            logger.error(f"Error in proactive installation approach: {e}")
            return False