import requests
import json
import logging
import msal  # You'll need to pip install msal

# Configure logging
logger = logging.getLogger(__name__)

class TeamsDirectMessaging:
    """Class to handle Microsoft Teams direct messaging via Graph API"""
    
    def __init__(self, client_id, client_secret, tenant_id):
        """Initialize with Azure AD credentials"""
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = None
        self.app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=f"https://login.microsoftonline.com/{tenant_id}"
        )
    
    def authenticate(self):
        """Authenticate with Microsoft Graph API using client credentials flow"""
        try:
            # Acquire token for application
            result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                logger.info("Successfully authenticated with Microsoft Graph API")
                return True
            else:
                logger.error(f"Authentication failed: {result.get('error_description', 'Unknown error')}")
                return False
        except Exception as e:
            logger.error(f"Error during authentication: {e}", exc_info=True)
            return False
    
    def get_user_id_by_email(self, email):
        """Get Teams user ID from email address"""
        if not self.access_token:
            if not self.authenticate():
                return None
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            response = requests.get(
                f"https://graph.microsoft.com/v1.0/users/{email}",
                headers=headers
            )
            
            if response.status_code == 200:
                user_data = response.json()
                return user_data.get("id")
            else:
                logger.error(f"Failed to get user ID for {email}. Status: {response.status_code}, Response: {response.text}")
                return None
        except Exception as e:
            logger.error(f"Error getting user ID: {e}", exc_info=True)
            return None
    
    def create_chat(self, user_id):
        """Create or get a 1:1 chat with the user"""
        if not self.access_token:
            if not self.authenticate():
                return None
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        # Create a 1:1 chat
        chat_data = {
            "chatType": "oneOnOne",
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_id}')"
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
                return chat_info.get("id")
            else:
                logger.error(f"Failed to create chat. Status: {response.status_code}, Response: {response.text}")
                return None
        except Exception as e:
            logger.error(f"Error creating chat: {e}", exc_info=True)
            return None
    
    def send_direct_message(self, chat_id, message_content):
        """Send a message to a Teams chat"""
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
            response = requests.post(
                f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages",
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
            logger.error(f"Error sending message: {e}", exc_info=True)
            return False

# Function to add to app.py
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
        max_days_overdue = max(t.get("Days Overdue", 0) for t in tasks)
        one_day = (max_days_overdue == 1)
        two_plus = (max_days_overdue >= 2)

        # Format title based on urgency
        if one_day:
            title = "Quick Nudge – Log Your Hours"
            emoji = "🟠"
            intro = ("This is a gentle reminder to log your hours for the "
                     "task below — it only takes a minute:")
        else:
            title = "Heads-Up: You've Missed Logging Hours for 2 Days"
            emoji = "🔴"
            intro = ("It looks like no hours have been logged for the past "
                     "two days for the task(s) below:")

        # Create HTML for tasks table
        tasks_html = "<table border='1' cellpadding='6' cellspacing='0'>"
        tasks_html += "<tr><th>Task</th><th>Project</th><th>Date</th><th>Client Success Contact</th></tr>"
        
        for t in tasks:
            tasks_html += f"""<tr>
                <td>{t.get('Task', 'Unknown')}</td>
                <td>{t.get('Project', 'Unknown')}</td>
                <td>{t.get('Date', '—')}</td>
                <td>{t.get('Client Success Member', 'Unknown')}</td>
            </tr>"""
        
        tasks_html += "</table>"

        # Create HTML message body
        message_html = f"""
        <h2>{emoji} {title}</h2>
        <p>Hi {designer_name},</p>
        <p>{intro}</p>
        
        {tasks_html}
        
        <p>
            {"Taking a minute now helps us stay on top of things later 🙌<br>Let us know if you need any support with this." 
            if one_day else 
            "We completely understand things can get busy — but consistent time logging helps us improve project planning and smooth reporting.<br>If something's holding you back from logging your hours, just reach out. We're here to help."}
        </p>
        
        <p style='font-size: 12px; color: gray;'>— Automated notice from the Missing Timesheet Reporter</p>
        """

        # Create or get chat with user
        chat_id = teams_client.create_chat(designer_teams_id)
        if not chat_id:
            logger.error(f"Failed to create chat with user {designer_name}")
            return False

        # Send message to chat
        message_sent = teams_client.send_direct_message(chat_id, message_html)
        
        if message_sent:
            logger.info(f"Direct message sent to {designer_name} via Teams")
            return True
        else:
            logger.error(f"Failed to send direct message to {designer_name}")
            return False

    except Exception as exc:
        logger.error(f"send_teams_direct_message failed: {exc}", exc_info=True)
        return False