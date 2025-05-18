# Add these imports if not already imported
import streamlit as st
import pandas as pd
import requests  # For making webhook HTTP requests
import xmlrpc.client
from datetime import datetime, timedelta, date
import logging
from io import BytesIO, StringIO
import traceback
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
import json
import msal  # You'll need to pip install msal
import time
from teams_direct_messaging import TeamsMessenger


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='planning_timesheet_reporter.log',
    force=True  # Add this line to reset any existing logging configuration
)
logger = logging.getLogger(__name__)

# Initialize session state
if 'odoo_uid' not in st.session_state:
    st.session_state.odoo_uid = None
if 'odoo_models' not in st.session_state:
    st.session_state.odoo_models = None
# Load secrets for Odoo connection
odoo_url = st.secrets.get("ODOO_URL", "")
odoo_db = st.secrets.get("ODOO_DB", "")
odoo_username = st.secrets.get("ODOO_USERNAME", "")
odoo_password = st.secrets.get("ODOO_PASSWORD", "")

if 'odoo_db' not in st.session_state:
    st.session_state.odoo_db = odoo_db
if 'odoo_url' not in st.session_state:
    st.session_state.odoo_url = odoo_url
if 'odoo_username' not in st.session_state:
    st.session_state.odoo_username = odoo_username
if 'odoo_password' not in st.session_state:
    st.session_state.odoo_password = odoo_password
if 'debug_mode' not in st.session_state:
    st.session_state.debug_mode = False
if 'confirmed_only' not in st.session_state:
    st.session_state.confirmed_only = True  # Default to showing only confirmed tasks
if 'shift_status_filter' not in st.session_state:
    st.session_state.shift_status_filter = "Planned"  # Default to Planned (confirmed)
if 'model_fields_cache' not in st.session_state:
    st.session_state.model_fields_cache = {}
if 'last_error' not in st.session_state:
    st.session_state.last_error = None
# Email settings
if 'email_enabled' not in st.session_state:
    st.session_state.email_enabled = True
if 'email_recipient' not in st.session_state:
    st.session_state.email_recipient = odoo_username  # Default to the Odoo username (usually an email)
if 'designer_emails_enabled' not in st.session_state:
    st.session_state.designer_emails_enabled = True
if 'designer_email_mapping' not in st.session_state:
    st.session_state.designer_email_mapping = {}
if 'smtp_server' not in st.session_state:
    st.session_state.smtp_server = "smtp.gmail.com"
# Add after the email settings initialization (around line 50-70)
if 'manager_emails_enabled' not in st.session_state:
    st.session_state.manager_emails_enabled = True
# Email settings
if 'smtp_port' not in st.session_state:
    st.session_state.smtp_port = 587
if 'smtp_username' not in st.session_state:
    st.session_state.smtp_username = ""
if 'smtp_password' not in st.session_state:
    st.session_state.smtp_password = ""

# Teams webhook settings
if 'webhooks_enabled' not in st.session_state:
    st.session_state.webhooks_enabled = False
if 'designer_webhook_mapping' not in st.session_state:
    st.session_state.designer_webhook_mapping = {}
if 'test_webhook_url' not in st.session_state:
    st.session_state.test_webhook_url = ""
# Teams direct messaging settings
if 'teams_direct_msg_enabled' not in st.session_state:
    st.session_state.teams_direct_msg_enabled = True
if 'azure_client_id' not in st.session_state:
    st.session_state.azure_client_id = ""
if 'azure_client_secret' not in st.session_state:
    st.session_state.azure_client_secret = ""
if 'azure_tenant_id' not in st.session_state:
    st.session_state.azure_tenant_id = ""
if 'designer_teams_id_mapping' not in st.session_state:
    st.session_state.designer_teams_id_mapping = {}


# Load Azure AD credentials from secrets if they exist
if hasattr(st.secrets, "AZURE_AD"):
    if "CLIENT_ID" in st.secrets.AZURE_AD:
        st.session_state.azure_client_id = st.secrets.AZURE_AD.CLIENT_ID
    if "CLIENT_SECRET" in st.secrets.AZURE_AD:
        st.session_state.azure_client_secret = st.secrets.AZURE_AD.CLIENT_SECRET
    if "TENANT_ID" in st.secrets.AZURE_AD:
        st.session_state.azure_tenant_id = st.secrets.AZURE_AD.TENANT_ID

# Load Teams user ID mappings from secrets if they exist
if hasattr(st.secrets, "TEAMS_USER_IDS"):
    for designer, teams_id in st.secrets.TEAMS_USER_IDS.items():
        st.session_state.designer_teams_id_mapping[designer] = teams_id
# Load webhook mappings from secrets if they exist
if hasattr(st.secrets, "WEBHOOKS"):
    for designer, webhook_url in st.secrets.WEBHOOKS.items():
        st.session_state.designer_webhook_mapping[designer] = webhook_url

# Load designer email mappings from secrets if they exist
if hasattr(st.secrets, "DESIGNER_EMAILS"):
    for designer, email in st.secrets.DESIGNER_EMAILS.items():
        st.session_state.designer_email_mapping[designer] = email

# Add reference date for cutoff of historical tasks
if 'reference_date' not in st.session_state:
    st.session_state.reference_date = date.today() - timedelta(days=7)  # Default to 7 days ago

# Add additional initialization for email settings from secrets if they exist
if hasattr(st.secrets, "EMAIL"):
    if "SMTP_SERVER" in st.secrets.EMAIL and st.session_state.smtp_server == "smtp.gmail.com":
        st.session_state.smtp_server = st.secrets.EMAIL.SMTP_SERVER
    if "SMTP_PORT" in st.secrets.EMAIL and st.session_state.smtp_port == 587:
        st.session_state.smtp_port = st.secrets.EMAIL.SMTP_PORT
    if "SMTP_USERNAME" in st.secrets.EMAIL and not st.session_state.smtp_username:
        st.session_state.smtp_username = st.secrets.EMAIL.SMTP_USERNAME
    if "SMTP_PASSWORD" in st.secrets.EMAIL and not st.session_state.smtp_password:
        st.session_state.smtp_password = st.secrets.EMAIL.SMTP_PASSWORD

if 'current_page' not in st.session_state:
    st.session_state.current_page = "Missing Timesheet Reporter"

# Initialize data container
if 'employee_data' not in st.session_state:
    st.session_state.employee_data = None

# Load employee data at startup
def load_employee_data():
    try:
        csv_path = "uhd_data.csv"
        if os.path.exists(csv_path):
            logger.info(f"Loading employee data from CSV: {csv_path}")
            df = pd.read_csv(csv_path)
            
            # Verify required columns exist
            required_columns = ["Employee Name", "Manager", "Work Email", "Microsoft ID"]
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                logger.error(f"Required columns missing from CSV: {missing_columns}")
                return None
            
            logger.info(f"Successfully loaded employee data with {len(df)} rows")
            return df
        else:
            logger.error(f"Employee data file not found: {csv_path}")
            return None
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error loading employee data: {e}\n{error_details}")
        return None
    
def update_designer_mappings_from_csv():
    """Update designer email and Teams ID mappings from CSV data"""
    if st.session_state.employee_data is None:
        logger.warning("Cannot update designer mappings: employee data not available")
        return
        
    df = st.session_state.employee_data
    
    # Clear existing mappings to prioritize CSV data
    st.session_state.designer_email_mapping = {}
    st.session_state.designer_teams_id_mapping = {}
    
    # Update designer email mappings
    for _, row in df.iterrows():
        if pd.notna(row["Employee Name"]) and pd.notna(row["Work Email"]):
            st.session_state.designer_email_mapping[row["Employee Name"]] = row["Work Email"]
    
    # Update Microsoft Teams ID mappings
    for _, row in df.iterrows():
        if pd.notna(row["Employee Name"]) and pd.notna(row["Microsoft ID"]):
            st.session_state.designer_teams_id_mapping[row["Employee Name"]] = row["Microsoft ID"]
                
    logger.info(f"Updated designer mappings from CSV data: {len(st.session_state.designer_email_mapping)} emails, {len(st.session_state.designer_teams_id_mapping)} Teams IDs")
# Load employee data at startup
st.session_state.employee_data = load_employee_data()
if st.session_state.employee_data is None:
    logger.warning("Failed to load employee data at startup. Some features may not work correctly.")

# Call update_designer_mappings_from_csv() after loading employee data
if st.session_state.employee_data is not None:
    update_designer_mappings_from_csv()


def send_designer_notification(designer_name, designer_teams_id, tasks):
    """Send a notification to a designer"""
    
    # Create messenger
    messenger = TeamsMessenger(
        st.session_state.azure_client_id,
        st.session_state.azure_client_secret,
        st.session_state.azure_tenant_id
    )
    
    # Format a compact message for topic
    max_days_overdue = max(t.get("Days Overdue", 0) for t in tasks)
    urgency_emoji = "🔴" if max_days_overdue >= 2 else "🟠"
    
    # Create a concise but informative topic
    task_summary = f"{len(tasks)} task{'s' if len(tasks) > 1 else ''}"
    oldest_date = min([t.get("Date", "") for t in tasks if t.get("Date")])
    
    # Format the notification topic
    message = f"{urgency_emoji} TIMESHEET ALERT - {task_summary} missing hours (oldest: {oldest_date}) - Action required"
    
    # Send notification
    return messenger.notify_user(designer_teams_id, message)

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
        st.info("You need to register an app in Azure AD with Microsoft Graph API permissions: Chat.Create, Chat.Read.All, Chat.ReadWrite.All, Teamwork.Migrate.All.")
        
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
        
        # Authentication test
        if st.button("Test Authentication"):
            if not (st.session_state.azure_client_id and st.session_state.azure_client_secret and st.session_state.azure_tenant_id):
                st.error("Please configure Azure AD credentials first")
            else:
                try:
                    # Create Teams messenger
                    messenger = TeamsMessenger(
                        st.session_state.azure_client_id,
                        st.session_state.azure_client_secret,
                        st.session_state.azure_tenant_id
                    )
                    
                    # Test authentication
                    with st.spinner("Testing authentication..."):
                        auth_result = messenger.authenticate()
                        if auth_result:
                            st.success("✅ Authentication successful!")
                        else:
                            st.error("❌ Authentication failed!")
                except Exception as e:
                    st.error(f"Error testing authentication: {str(e)}")
        
        # Designer to Teams ID mapping
        st.markdown("### Designer Teams User ID Mapping")
        
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
                        st.rerun()
        
        # Test message section
        st.markdown("### Test Message")
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
                
                # Create test task
                test_task = [{
                    "Project": "Test Project",
                    "Task": "Test Task",
                    "Start Time": "09:00",
                    "End Time": "17:00",
                    "Allocated Hours": 8.0,
                    "Date": time.strftime("%Y-%m-%d"),
                    "Days Overdue": 1,
                    "Client Success Member": "Test Manager"
                }]
                
                with st.spinner("Sending test message..."):
                    # Send test notification
                    message_sent = send_designer_notification(
                        test_designer,
                        teams_id,
                        test_task
                    )
                    
                    if message_sent:
                        st.success(f"Message sent to {test_designer}! Check your Teams app.")
                    else:
                        st.error(f"Failed to send message to {test_designer}")
def send_designer_teams_direct_messages(designers, selected_date):
    """Send Teams direct messages to designers with missing timesheets"""
    if not st.session_state.teams_direct_msg_enabled:
        return False, 0, 0
        
    if not (st.session_state.azure_client_id and st.session_state.azure_client_secret and st.session_state.azure_tenant_id):
        return False, 0, 0
    
    success_count = 0
    fail_count = 0
    
    # Send direct message to each designer with missing timesheets
    for designer, tasks in designers.items():
        # Check if we have a Teams ID for this designer
        if designer in st.session_state.designer_teams_id_mapping:
            designer_teams_id = st.session_state.designer_teams_id_mapping[designer]
            
            # Send the notification
            message_sent = send_designer_notification(
                designer,
                designer_teams_id,
                tasks
            )
            
            if message_sent:
                success_count += 1
            else:
                fail_count += 1
        else:
            fail_count += 1
    
    return True, success_count, fail_count

def authenticate_odoo(url, db, username, password):
    """Authenticate with Odoo and return uid and models"""
    try:
        common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common")
        uid = common.authenticate(db, username, password, {})
        
        if not uid:
            st.error("Odoo authentication failed - invalid credentials")
            logger.error("Odoo authentication failed - invalid credentials")
            return None, None
            
        models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object")
        logger.info(f"Successfully connected to Odoo (UID: {uid})")
        return uid, models
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Odoo connection error: {e}\n{error_details}")
        st.error(f"Odoo connection error: {e}")
        st.session_state.last_error = error_details
        return None, None

def get_model_fields(models, uid, odoo_db, odoo_password, model_name):
    """Get fields for a specific model, with caching"""
    # Check if we have cached fields for this model
    if model_name in st.session_state.model_fields_cache:
        return st.session_state.model_fields_cache[model_name]
    
    try:
        fields = models.execute_kw(
            odoo_db, uid, odoo_password,
            model_name, 'fields_get',
            [],
            {'attributes': ['string', 'type', 'relation']}
        )
        # Cache the result
        st.session_state.model_fields_cache[model_name] = fields
        return fields
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error getting fields for model {model_name}: {e}\n{error_details}")
        st.session_state.last_error = error_details
        return {}

def get_planning_slots(models, uid, odoo_db, odoo_password, start_date, end_date=None, shift_status_filter=None):
    """
    Get planning slots for a date range, with a focus on finding all slots 
    that overlap with the given date range. Optionally filter by x_studio_shift_status.
    """
    try:
        # Get the fields for planning.slot model
        fields_info = get_model_fields(models, uid, odoo_db, odoo_password, 'planning.slot')
        available_fields = list(fields_info.keys())
        
        # Handle single date or date range
        if end_date is None:
            end_date = start_date
        
        # Prepare the date strings
        start_date_str = start_date.strftime("%Y-%m-%d")
        end_date_str = end_date.strftime("%Y-%m-%d")
        next_date_str = (end_date + timedelta(days=1)).strftime("%Y-%m-%d")
        
        # Create different domain variations to catch various date formats and conditions
        # This is needed because Odoo instances can have different field naming or formats
        base_domains = [
            # Standard format with start/end datetime - get all slots in the date range
            [
                '|',
                # Slots that start in our date range
                '&', ('start_datetime', '>=', f"{start_date_str} 00:00:00"), ('start_datetime', '<', f"{next_date_str} 00:00:00"),
                # Slots that end in our date range or overlap with it
                '|',
                '&', ('end_datetime', '>=', f"{start_date_str} 00:00:00"), ('end_datetime', '<', f"{next_date_str} 00:00:00"),
                '&', ('start_datetime', '<', f"{start_date_str} 00:00:00"), ('end_datetime', '>=', f"{next_date_str} 00:00:00")
            ],
            # Alternative based on date fields if they exist
            [],
            # Simple date string matching (fallback)
            [('start_datetime', '>=', start_date_str), ('start_datetime', '<', next_date_str)]
        ]
        
        # Add shift_status filter if provided
        domains = []
        if shift_status_filter and 'x_studio_shift_status' in available_fields:
            logger.info(f"Filtering planning slots by x_studio_shift_status: {shift_status_filter}")
            for base_domain in base_domains:
                if base_domain:  # Skip empty domain
                    domain_with_filter = base_domain.copy()
                    domain_with_filter.append(('x_studio_shift_status', '=', shift_status_filter))
                    domains.append(domain_with_filter)
        else:
            domains = base_domains
        
        # Basic fields we want, checking which ones exist
        desired_fields = [
            'id', 'name', 'resource_id', 'start_datetime', 'end_datetime', 
            'allocated_hours', 'state', 'project_id', 'task_id', 'x_studio_shift_status',
            'create_uid', 'x_studio_sub_task_1', 'x_studio_task_activity', 'x_studio_service_category_1', 'x_studio_sub_task_link' 
        ]
        
        # Only request fields that exist
        fields_to_request = [f for f in desired_fields if f in available_fields]
        
        # Log the fields we're requesting
        logger.info(f"Requesting planning slot fields: {fields_to_request}")
        
        # Try each domain until we get results
        all_slots = []
        success = False
        
        for i, domain in enumerate(domains):
            if not domain:  # Skip empty domain
                continue
                
            try:
                logger.info(f"Trying planning slot domain {i+1}: {domain}")
                slots = models.execute_kw(
                    odoo_db, uid, odoo_password,
                    'planning.slot', 'search_read',
                    [domain],
                    {'fields': fields_to_request}
                )
                
                if slots:
                    logger.info(f"Found {len(slots)} planning slots with domain {i+1}")
                    all_slots.extend(slots)
                    success = True
                    # Don't break, try all domains to get comprehensive results
            except Exception as e:
                # Just log and continue to next domain
                logger.warning(f"Error with planning slot domain {i+1}: {e}")
        
        # If we didn't get any results, try a more permissive approach
        if not success:
            try:
                logger.info("Trying to get all recent planning slots")
                # Get all slots from recent dates
                one_month_ago = (start_date - timedelta(days=30)).strftime("%Y-%m-%d")
                base_domain = [('start_datetime', '>=', one_month_ago)]
                        
                # Add shift_status filter if provided
                if shift_status_filter and 'x_studio_shift_status' in available_fields:
                    base_domain.append(('x_studio_shift_status', '=', shift_status_filter))
                
                recent_slots = models.execute_kw(
                    odoo_db, uid, odoo_password,
                    'planning.slot', 'search_read',
                    [base_domain],
                    {'fields': fields_to_request}
                )
                
                # Filter by date string to find matching ones
                end_date_str_simple = end_date_str.replace('-', '')  # Also try without dashes
                
                for slot in recent_slots:
                    start = slot.get('start_datetime', '')
                    if end_date_str in start or end_date_str_simple in start.replace('-', ''):
                        all_slots.append(slot)
                
                logger.info(f"Filtered to {len(all_slots)} planning slots for the date range")
                
            except Exception as e:
                error_details = traceback.format_exc()
                logger.error(f"Error with permissive planning slot query: {e}\n{error_details}")
        
        # Deduplicate slots by ID
        unique_slots = []
        seen_ids = set()
        for slot in all_slots:
            if slot['id'] not in seen_ids:
                unique_slots.append(slot)
                seen_ids.add(slot['id'])
        
        logger.info(f"Returning {len(unique_slots)} unique planning slots for date range {start_date_str} to {end_date_str}")
        return unique_slots
        
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error fetching planning slots: {e}\n{error_details}")
        st.error(f"Error fetching planning slots: {e}")
        st.session_state.last_error = error_details
        return []

def get_timesheet_entries(models, uid, odoo_db, odoo_password, start_date, end_date=None):
    """Get timesheet entries for a date range"""
    try:
        if end_date is None:
            end_date = start_date
            
        # Add one day to end_date to include the entire end date
        query_end_date = end_date + timedelta(days=1)
            
        start_date_str = start_date.strftime("%Y-%m-%d")
        end_date_str = query_end_date.strftime("%Y-%m-%d")
        
        # Domain for date range
        domain = [
            ('date', '>=', start_date_str),
            ('date', '<', end_date_str)
        ]
        
        # Get fields for the model to make sure we only request valid fields
        fields_info = get_model_fields(models, uid, odoo_db, odoo_password, 'account.analytic.line')
        available_fields = list(fields_info.keys())
        
        # Fields we want (if they exist)
        desired_fields = [
            'id', 'name', 'date', 'unit_amount', 'employee_id', 
            'task_id', 'project_id', 'user_id', 'company_id'
        ]
        
        # Only request fields that exist
        fields_to_request = [f for f in desired_fields if f in available_fields]
        
        # Execute query
        logger.info(f"Querying timesheets with domain: {domain}")
        entries = models.execute_kw(
            odoo_db, uid, odoo_password,
            'account.analytic.line', 'search_read',
            [domain],
            {'fields': fields_to_request}
        )
        
        logger.info(f"Found {len(entries)} timesheet entries")
        return entries
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error fetching timesheet entries: {e}\n{error_details}")
        st.error(f"Error fetching timesheet entries: {e}")
        st.session_state.last_error = error_details
        return []

def get_references_data(models, uid, odoo_db, odoo_password):
    """Get reference data (projects, users, employees, etc.) for display"""
    reference_data = {}
    
    try:
        # Get resources (employees/equipment in planning)
        resources = models.execute_kw(
            odoo_db, uid, odoo_password,
            'resource.resource', 'search_read',
            [[]],
            {'fields': ['id', 'name', 'user_id', 'resource_type', 'company_id']}
        )
        reference_data['resources'] = {r['id']: r for r in resources}
        
        # Get projects
        projects = models.execute_kw(
            odoo_db, uid, odoo_password,
            'project.project', 'search_read',
            [[]],
            {'fields': ['id', 'name']}
        )
        reference_data['projects'] = {p['id']: p for p in projects}
        
        # Get users
        users = models.execute_kw(
            odoo_db, uid, odoo_password,
            'res.users', 'search_read',
            [[]],
            {'fields': ['id', 'name']}
        )
        reference_data['users'] = {u['id']: u for u in users}
        
        # Get tasks
        tasks = models.execute_kw(
            odoo_db, uid, odoo_password,
            'project.task', 'search_read',
            [[]],
            {'fields': ['id', 'name']}
        )
        reference_data['tasks'] = {t['id']: t for t in tasks}
        
        return reference_data
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error fetching reference data: {e}\n{error_details}")
        st.warning(f"Error fetching some reference data: {e}")
        st.session_state.last_error = error_details
        return reference_data

def send_email_report(df, selected_date, missing_count, timesheet_count, shift_status_filter=None, reference_date=None):
    """Send email with report attached as CSV and summary in the body"""
    try:
        if not st.session_state.email_enabled:
            logger.info("Email sending is disabled, skipping")
            return False
            
        if not st.session_state.smtp_username or not st.session_state.smtp_password:
            logger.error("Email credentials not configured")
            return False
            
        # Create email
        msg = MIMEMultipart()
        msg['From'] = st.session_state.smtp_username
        msg['To'] = st.session_state.email_recipient
        
        # Create date strings for display
        if reference_date:
            msg['Subject'] = f"Missing Timesheet Report - {reference_date.strftime('%Y-%m-%d')} to {selected_date.strftime('%Y-%m-%d')}"
            date_range_str = f"{reference_date.strftime('%Y-%m-%d')} to {selected_date.strftime('%Y-%m-%d')}"
        else:
            msg['Subject'] = f"Missing Timesheet Report - {selected_date.strftime('%Y-%m-%d')}"
            date_range_str = selected_date.strftime('%Y-%m-%d')
        
        # Prepare filter info for email body
        filter_text = ""
        if shift_status_filter:
            filter_text = f" with shift status '{shift_status_filter}'"
            
        # Create email body text
        body = f"""
        <html>
        <body>
        <h2>Missing Timesheet Report - {date_range_str}</h2>
        <p>This is an automated report from the Missing Timesheet Reporter tool.</p>
        
        <h3>Summary:</h3>
        <ul>
            <li>Date Range: {date_range_str}</li>
            <li>Found {timesheet_count} timesheet entries</li>
            <li>Found {missing_count} planning slots{filter_text} without timesheet entries</li>
        </ul>
        """
        
        # Add empty data message or summary of missing entries
        if df.empty:
            body += f"<p>No planning slots found for this date range{filter_text}.</p>"
        elif missing_count == 0:
            body += f"<p>All planning slots{filter_text} have corresponding timesheet entries!</p>"
        else:
            # Add summary table of designers with missing entries
            designer_summary = df.groupby("Designer").size().reset_index(name="Missing Entries")
            designer_summary = designer_summary.sort_values("Missing Entries", ascending=False)
            
            body += "<h3>Missing Entries by Designer:</h3><table border='1'><tr><th>Designer</th><th>Missing Entries</th></tr>"
            for _, row in designer_summary.iterrows():
                body += f"<tr><td>{row['Designer']}</td><td>{row['Missing Entries']}</td></tr>"
            body += "</table>"
            
            # Add summary table of projects with missing entries
            project_summary = df.groupby("Project").size().reset_index(name="Missing Entries")
            project_summary = project_summary.sort_values("Missing Entries", ascending=False)
            
            body += "<h3>Missing Entries by Project:</h3><table border='1'><tr><th>Project</th><th>Missing Entries</th></tr>"
            for _, row in project_summary.iterrows():
                body += f"<tr><td>{row['Project']}</td><td>{row['Missing Entries']}</td></tr>"
            body += "</table>"
            
        body += """
        <p>Please check the attached CSV file for detailed information.</p>
        <p>This is an automated message from the Missing Timesheet Reporter tool.</p>
        </body>
        </html>
        """
        
        # Attach email body
        msg.attach(MIMEText(body, 'html'))
        
        # Create CSV attachment if data exists
        if not df.empty and missing_count > 0:
            csv_data = df.to_csv(index=False)
            attachment = MIMEApplication(csv_data.encode('utf-8'))
            date_id = selected_date.strftime("%Y-%m-%d")
            attachment['Content-Disposition'] = f'attachment; filename="missing_timesheet_report_{date_id}.csv"'
            msg.attach(attachment)
            
        # Send email
        server = smtplib.SMTP(st.session_state.smtp_server, st.session_state.smtp_port)
        server.starttls()
        server.login(st.session_state.smtp_username, st.session_state.smtp_password)
        server.send_message(msg)
        server.quit()
        
        logger.info(f"Email report sent to {st.session_state.email_recipient}")
        return True
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error sending email: {e}\n{error_details}")
        st.session_state.last_error = error_details
        return False

def send_designer_email(
        designer_name: str,
        designer_email: str,
        report_date: date,
        tasks: list,
        smtp_settings: dict
):
    """
    Send a one-shot e-mail to a designer listing the tasks that still
    have no hours logged.  Message style changes according to how long
    the oldest task has been open.
    """
    try:
        # -- SMTP sanity check ------------------------------------------------
        for key in ("server", "port", "username", "password"):
            if key not in smtp_settings or not smtp_settings[key]:
                logger.error("Missing SMTP setting: %s", key)
                return False

        # -- work out 'days overdue' -----------------------------------------
        max_days_overdue = max(t.get("Days Overdue", 0) for t in tasks)

        one_day  = (max_days_overdue == 1)
        two_plus = (max_days_overdue >= 2)
        
        # -- Get the first name only -----------------------------------------
        # Extract first name from the full name
        first_name = designer_name.split()[0] if designer_name else "there"

        if one_day:
            subj = "Quick Nudge – Log Your Hours"
        else:  # two_plus
            subj = "Heads-Up: You've Missed Logging Hours for 2 Days"

        # -- e-mail boilerplate ----------------------------------------------
        msg            = MIMEMultipart()
        msg["From"]    = smtp_settings["username"]
        msg["To"]      = designer_email
        msg["Subject"] = subj

        # --------------------------------------------------------------------
        # Build HTML body
        # --------------------------------------------------------------------
        def format_task(t):
            task_name = t.get('Task', 'Unknown')
            task_link = t.get('Sub_Task_Link', '')
            
            # Format the link as a separate column
            link_display = "No link"
            if task_link and isinstance(task_link, str):
                if task_link.startswith('http') or task_link.startswith('/'):
                    link_display = f'<a href="{task_link}" target="_blank">Open Task</a>'
            
            return f"""
            <tr>
                <td>{task_name}</td>
                <td>{t.get('Project', 'Unknown')}</td>
                <td>{t.get('Date', '—')}</td>
                <td>{t.get('Client Success Member', 'Unknown')}</td>
                <td>{link_display}</td>
            </tr>"""

        tasks_html = "".join(format_task(t) for t in tasks)

        # Use first name instead of full name
        greeting = (
            f"<p>Hi {first_name},</p>"
            if one_day
            else f"<p>Hi {first_name},</p>"
        )

        intro = ("""
            This is a gentle reminder to log your hours for the task below — 
            it takes a minute, but the impact is big:
        """ if one_day else """
            It looks like no hours have been logged for the past two days
            for the following task:
        """)

        outro = ("""
            <p>Taking a minute now helps us stay on top of things later 🙌</p>
            <p>Let us know if you need any support with this.</p>
        """ if one_day else """
            <p>We completely understand things can get busy — but consistent
            time logging helps us improve project planning and smooth
            reporting.</p>
            <p>If something's holding you back from logging your hours,
            just reach out. We're here to help.</p>
        """)

        body = f"""
        <html><body>
        {greeting}
        <p>{intro}</p>

        <table border="1" cellpadding="6" cellspacing="0">
            <thead>
                <tr>
                    <th>Task</th>
                    <th>Project</th>
                    <th>Assigned on / Dates</th>
                    <th>Client-Success Contact</th>
                    <th>Link</th>
                </tr>
            </thead>
            <tbody>
                {tasks_html}
            </tbody>
        </table>

        {outro}

        <p style="font-size: 12px;">
            — Automated notice from the Missing Timesheet Reporter
        </p>
        </body></html>
        """

        msg.attach(MIMEText(body, "html"))

        # -- send it ----------------------------------------------------------
        server = smtplib.SMTP(smtp_settings["server"], smtp_settings["port"])
        server.starttls()
        server.login(smtp_settings["username"], smtp_settings["password"])
        server.send_message(msg)
        server.quit()

        logger.info("Designer e-mail sent to %s", designer_email)
        return True

    except Exception as exc:
        logger.error("send_designer_email failed: %s", exc, exc_info=True)
        return False

# Add these functions after send_designer_email function (around line 487)
def load_employee_manager_mapping():
    """Load employee-manager relationships from preloaded employee data"""
    try:
        if st.session_state.employee_data is None:
            # Try loading again if not already loaded
            st.session_state.employee_data = load_employee_data()
            
        if st.session_state.employee_data is None:
            logger.error("Cannot load employee-manager mapping: employee data not available")
            return {}
            
        df = st.session_state.employee_data
        
        # Process each employee row
        mapping = {}
        for _, row in df.iterrows():
            try:
                employee_name = row["Employee Name"]
                manager_name = row["Manager"]
                work_email = row["Work Email"]
                
                if pd.notna(employee_name) and pd.notna(manager_name):
                    # Find manager's email by looking up the manager in the dataframe
                    manager_row = df[df["Employee Name"] == manager_name]
                    if not manager_row.empty and pd.notna(manager_row.iloc[0]["Work Email"]):
                        manager_email = manager_row.iloc[0]["Work Email"]
                    else:
                        logger.warning(f"Could not find email for manager '{manager_name}' of employee '{employee_name}'")
                        continue
                    
                    # Store the mapping with the employee name as the key
                    mapping[employee_name] = {
                        "manager_name": manager_name,
                        "manager_email": manager_email
                    }
            except Exception as e:
                logger.warning(f"Error processing row for employee {row.get('Employee Name', 'Unknown')}: {e}")
                continue
        
        logger.info(f"Loaded {len(mapping)} employee-manager relationships")
        return mapping
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error loading employee mapping: {e}\n{error_details}")
        return {}
    
def send_manager_email(manager_name, manager_email, designers_tasks, selected_date):
    """Send email to a manager about their team's missing timesheets"""
    try:
        logger.info(f"Preparing email for manager: {manager_name} ({manager_email})")
        
        # Get the first name only
        first_name = manager_name.split()[0] if manager_name else "there"
        
        # Determine if any tasks are more than 1 day overdue
        max_days_overdue = 0
        for designer_tasks in designers_tasks.values():
            for task in designer_tasks:
                days_overdue = task.get("Days Overdue", 0)
                max_days_overdue = max(max_days_overdue, days_overdue)
        
        logger.info(f"Maximum days overdue for any task: {max_days_overdue}")
        
        # Set email subject and greeting based on days overdue
        if max_days_overdue >= 2:
            subject = f"Urgent: Team Members Haven't Logged Hours for 2 Days"
            greeting = f"Hi {first_name},"
            intro_text = "We've noticed that the following team members have <b>not logged their hours for 2 consecutive days</b> on assigned tasks. This is creating delays in tracking and reporting:"
            closing = "This needs immediate follow-up. Please address this with your team and make sure all pending hours are logged without further delay.\n\nLet us know if any blockers are preventing this from happening."
        else:
            subject = f"Unlogged Hours Report – {selected_date.strftime('%Y-%m-%d')}"
            greeting = f"Hi {first_name},"
            intro_text = f"The following team members haven't logged their hours for tasks assigned on <b>{selected_date.strftime('%Y-%m-%d')}</b>:"
            closing = "Reminders have already been sent to the individuals. Kindly follow up as needed to ensure all hours are logged promptly.\n\nLet us know if you need anything else."
        
        logger.info(f"Email subject: {subject}")
        
        # Create email
        msg = MIMEMultipart('alternative')
        msg['From'] = st.session_state.smtp_username
        msg['To'] = manager_email
        msg['Subject'] = subject
        
        # Log the SMTP details being used
        logger.info(f"Using SMTP server: {st.session_state.smtp_server}:{st.session_state.smtp_port}")
        logger.info(f"Using SMTP username: {st.session_state.smtp_username}")
        
        # Email content creation logic remains the same
        html_body = f"""
        <html>
        <body>
        <p>{greeting}</p>
        <p>{intro_text}</p>
        <ol>
        """
        
        # Add each designer and their tasks in HTML
        designer_counter = 1
        for designer_name, tasks in designers_tasks.items():
            # Sort tasks by days overdue (descending)
            sorted_tasks = sorted(tasks, key=lambda x: x.get('Days Overdue', 0), reverse=True)
            
            html_body += f"<li><b>{designer_name}</b>\n<ul>\n"
            
            for task in sorted_tasks:
                # Add task details
                html_body += f"<li><b>Project</b>: {task.get('Project', 'Unknown')}</li>\n"
                html_body += f"<li><b>Task</b>: {task.get('Task', 'Unknown')}</li>\n"
                
                if max_days_overdue >= 2:
                    html_body += f"<li><b>Assignment Dates</b>: {task.get('Date', 'Unknown')}</li>\n"
                else:
                    html_body += f"<li><b>Time Assigned</b>: {task.get('Start Time', 'Unknown')}</li>\n"
                
                html_body += f"<li><b>Client Success Contact</b>: {task.get('Client Success Member', 'Unknown')}</li>\n"
            
            html_body += "</ul></li>\n"
            designer_counter += 1
        
        # Add closing in HTML
        html_body += f"""
        </ol>
        <p>{closing}</p>
        <p>Thanks,<br>— Operations Team</p>
        </body>
        </html>
        """
        
        # Create plain text version
        text_body = f"{greeting}\n\n{intro_text}\n\n"
        
        # Add each designer and their tasks
        designer_counter = 1
        for designer_name, tasks in designers_tasks.items():
            # Sort tasks by days overdue (descending)
            sorted_tasks = sorted(tasks, key=lambda x: x.get('Days Overdue', 0), reverse=True)
            
            text_body += f"{designer_counter}. {designer_name}\n"
            
            for task in sorted_tasks:
                # Add task details
                text_body += f"  • Project: {task.get('Project', 'Unknown')}\n"
                text_body += f"  • Task: {task.get('Task', 'Unknown')}\n"
                
                if max_days_overdue >= 2:
                    text_body += f"  • Assignment Dates: {task.get('Date', 'Unknown')}\n"
                else:
                    text_body += f"  • Time Assigned: {task.get('Start Time', 'Unknown')}\n"
                
                text_body += f"  • Client Success Contact: {task.get('Client Success Member', 'Unknown')}\n"
            
            text_body += "\n"
            designer_counter += 1
        
        # Add closing
        text_body += f"{closing}\n\nThanks,\n— Operations Team"
        
        # Attach both versions
        msg.attach(MIMEText(text_body, 'plain'))
        msg.attach(MIMEText(html_body, 'html'))
        
        # Send email with detailed error handling
        try:
            logger.info(f"Connecting to SMTP server {st.session_state.smtp_server}:{st.session_state.smtp_port}")
            server = smtplib.SMTP(st.session_state.smtp_server, st.session_state.smtp_port)
            server.set_debuglevel(1)  # Enable SMTP debug output
            
            logger.info("Starting TLS")
            server.starttls()
            
            logger.info(f"Logging in with username: {st.session_state.smtp_username}")
            server.login(st.session_state.smtp_username, st.session_state.smtp_password)
            
            logger.info(f"Sending email to: {manager_email}")
            server.send_message(msg)
            
            logger.info("Quitting SMTP connection")
            server.quit()
            
            logger.info(f"Manager notification sent to {manager_name} ({manager_email})")
            return True
        except smtplib.SMTPAuthenticationError as e:
            logger.error(f"SMTP Authentication failed: {e}")
            return False
        except smtplib.SMTPException as e:
            logger.error(f"SMTP error: {e}")
            return False
        except Exception as e:
            logger.error(f"Unknown error sending email: {e}")
            return False
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error preparing manager email: {e}\n{error_details}")
        return False

def send_manager_notifications(designers, selected_date, reference_date=None):
    """Send email notifications to managers about their team members' missing timesheets"""
    logger.info(f"Starting manager notifications for {len(designers)} designers")
    
    if not st.session_state.email_enabled:
        logger.info("Email sending is disabled, skipping manager notifications")
        return False, 0, 0
        
    if not (st.session_state.smtp_server and 
            st.session_state.smtp_port and
            st.session_state.smtp_username and 
            st.session_state.smtp_password):
        logger.error("Email settings not configured")
        return False, 0, 0
    
    try:
        # Load employee-manager mapping
        employee_manager_mapping = load_employee_manager_mapping()
        if not employee_manager_mapping:
            logger.error("Could not load employee-manager mapping")
            return False, 0, 0
        
        # Group tasks by manager
        managers_tasks = {}
        
        # Debug log all designers
        logger.info(f"Designers with missing timesheets: {list(designers.keys())}")
        
        for designer, tasks in designers.items():
            logger.info(f"Processing designer: '{designer}' with {len(tasks)} missing tasks")
            
            # Find the manager for this designer
            if designer in employee_manager_mapping:
                manager_info = employee_manager_mapping[designer]
                manager_name = manager_info["manager_name"]
                manager_email = manager_info["manager_email"]
                
                logger.info(f"Found manager for '{designer}': {manager_name} ({manager_email})")
                
                if manager_name not in managers_tasks:
                    managers_tasks[manager_name] = {
                        "email": manager_email,
                        "designers": {}
                    }
                
                if designer not in managers_tasks[manager_name]["designers"]:
                    managers_tasks[manager_name]["designers"][designer] = []
                
                managers_tasks[manager_name]["designers"][designer].extend(tasks)
                logger.info(f"Added {len(tasks)} tasks for '{designer}' to manager '{manager_name}'")
            else:
                logger.warning(f"No manager found for designer '{designer}' - check name spelling in data file")
        
        # Send emails to managers
        success_count = 0
        fail_count = 0
        
        logger.info(f"Preparing to send emails to {len(managers_tasks)} managers")
        
        for manager_name, manager_data in managers_tasks.items():
            logger.info(f"Sending email to manager: {manager_name} ({manager_data['email']})")
            logger.info(f"Email will include {len(manager_data['designers'])} team members with missing timesheets")
            
            # List the team members for debugging
            for designer_name, designer_tasks in manager_data["designers"].items():
                logger.info(f"  - Designer: {designer_name} has {len(designer_tasks)} missing entries")
            
            # Prepare email content
            email_sent = send_manager_email(
                manager_name,
                manager_data["email"],
                manager_data["designers"],
                selected_date
            )
            
            if email_sent:
                logger.info(f"Successfully sent email to manager: {manager_name}")
                success_count += 1
            else:
                logger.error(f"Failed to send email to manager: {manager_name}")
                fail_count += 1
        
        logger.info(f"Manager notification summary: {success_count} successful, {fail_count} failed")
        return True, success_count, fail_count
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error sending manager notifications: {e}\n{error_details}")
        return False, 0, 0

def generate_missing_timesheet_report(selected_date, shift_status_filter=None, send_email=False, send_designer_emails=False):
    """
    Generate report of planning slots without timesheet entries for a date range from reference_date to selected_date
    """
    uid = st.session_state.odoo_uid
    models = st.session_state.odoo_models
    odoo_db = st.session_state.odoo_db
    odoo_password = st.session_state.odoo_password
    reference_date = st.session_state.reference_date
    
    if not uid or not models:
        st.error("Not connected to Odoo")
        return pd.DataFrame(), 0, 0
    
    try:
        # Step 1: Get all planning slots for the date range (reference_date to selected_date) with optional shift status filter
        planning_slots = get_planning_slots(models, uid, odoo_db, odoo_password, reference_date, selected_date, shift_status_filter)
        
        # Post-process to ensure only slots with the correct shift status are included
        # This adds a second layer of filtering in case the Odoo query didn't filter properly
        if shift_status_filter:
            filtered_slots = []
            for slot in planning_slots:
                slot_shift_status = slot.get('x_studio_shift_status', '')
                if slot_shift_status == shift_status_filter:
                    filtered_slots.append(slot)
            planning_slots = filtered_slots
            logger.info(f"Post-filtered to {len(planning_slots)} slots with x_studio_shift_status={shift_status_filter}")
        
        # Step 2: Get all timesheet entries for the date range
        timesheet_entries = get_timesheet_entries(models, uid, odoo_db, odoo_password, reference_date, selected_date)
        
        # Step 3: Get reference data
        ref_data = get_references_data(models, uid, odoo_db, odoo_password)
        
        # Step 4: Create resource+task+project to timesheet entry mapping
        # This ensures we match timesheets to specific tasks, not just to designers
        resource_task_to_timesheet = {}
        for entry in timesheet_entries:
            employee_id = None
            task_id = None
            project_id = None
            
            # Get employee ID
            if 'employee_id' in entry and entry['employee_id']:
                if isinstance(entry['employee_id'], list):
                    employee_id = entry['employee_id'][0]
                elif isinstance(entry['employee_id'], int):
                    employee_id = entry['employee_id']
            
            # Get task ID 
            if 'task_id' in entry and entry['task_id']:
                if isinstance(entry['task_id'], list):
                    task_id = entry['task_id'][0]
                elif isinstance(entry['task_id'], int):
                    task_id = entry['task_id']
            
            # Get project ID
            if 'project_id' in entry and entry['project_id']:
                if isinstance(entry['project_id'], list):
                    project_id = entry['project_id'][0]
                elif isinstance(entry['project_id'], int):
                    project_id = entry['project_id']
            
            # Get user ID (who actually created/logged the entry)
            user_id = None
            if 'user_id' in entry and entry['user_id']:
                if isinstance(entry['user_id'], list):
                    user_id = entry['user_id'][0]
                elif isinstance(entry['user_id'], int):
                    user_id = entry['user_id']
            
            if employee_id:
                # Create a unique key combining resource, task, and project
                # If task or project is None, we'll still create a key
                key = (employee_id, task_id, project_id)
                
                if key in resource_task_to_timesheet:
                    resource_task_to_timesheet[key]['hours'] += entry.get('unit_amount', 0)
                    resource_task_to_timesheet[key]['entries'].append(entry)
                    resource_task_to_timesheet[key]['user_ids'].add(user_id)
                else:
                    resource_task_to_timesheet[key] = {
                        'hours': entry.get('unit_amount', 0),
                        'entries': [entry],
                        'user_ids': {user_id} if user_id else set()
                    }
        
        # Add name-based mapping as a fallback
        # Sometimes the IDs don't match correctly but names will
        designer_name_to_timesheet = {}
        for entry in timesheet_entries:
            employee_name = None
            task_id = None
            project_id = None
            
            # Get employee name
            if 'employee_id' in entry and entry['employee_id'] and isinstance(entry['employee_id'], list) and len(entry['employee_id']) > 1:
                employee_name = entry['employee_id'][1]  # The second element in many Odoo relations is the name
            
            # Get task ID 
            if 'task_id' in entry and entry['task_id']:
                if isinstance(entry['task_id'], list):
                    task_id = entry['task_id'][0]
                elif isinstance(entry['task_id'], int):
                    task_id = entry['task_id']
            
            # Get project ID
            if 'project_id' in entry and entry['project_id']:
                if isinstance(entry['project_id'], list):
                    project_id = entry['project_id'][0]
                elif isinstance(entry['project_id'], int):
                    project_id = entry['project_id']
            
            # Get user ID (who actually created/logged the entry)
            user_id = None
            if 'user_id' in entry and entry['user_id']:
                if isinstance(entry['user_id'], list):
                    user_id = entry['user_id'][0]
                elif isinstance(entry['user_id'], int):
                    user_id = entry['user_id']
            
            if employee_name:
                # Create a unique key combining employee name, task, and project
                key = (employee_name, task_id, project_id)
                
                if key in designer_name_to_timesheet:
                    designer_name_to_timesheet[key]['hours'] += entry.get('unit_amount', 0)
                    designer_name_to_timesheet[key]['entries'].append(entry)
                    designer_name_to_timesheet[key]['user_ids'].add(user_id)
                else:
                    designer_name_to_timesheet[key] = {
                        'hours': entry.get('unit_amount', 0),
                        'entries': [entry],
                        'user_ids': {user_id} if user_id else set()
                    }
        
        # Also create a name-only mapping as a last resort
        designer_name_only_to_timesheet = {}
        for entry in timesheet_entries:
            employee_name = None
            
            # Get employee name
            if 'employee_id' in entry and entry['employee_id'] and isinstance(entry['employee_id'], list) and len(entry['employee_id']) > 1:
                employee_name = entry['employee_id'][1]
            
            # Get user ID (who actually created/logged the entry)
            user_id = None
            if 'user_id' in entry and entry['user_id']:
                if isinstance(entry['user_id'], list):
                    user_id = entry['user_id'][0]
                elif isinstance(entry['user_id'], int):
                    user_id = entry['user_id']
            
            if employee_name:
                if employee_name in designer_name_only_to_timesheet:
                    designer_name_only_to_timesheet[employee_name]['hours'] += entry.get('unit_amount', 0)
                    designer_name_only_to_timesheet[employee_name]['entries'].append(entry)
                    designer_name_only_to_timesheet[employee_name]['user_ids'].add(user_id)
                else:
                    designer_name_only_to_timesheet[employee_name] = {
                        'hours': entry.get('unit_amount', 0),
                        'entries': [entry],
                        'user_ids': {user_id} if user_id else set()
                    }
        
        # Log the mappings to help with debugging
        logger.info(f"Found {len(resource_task_to_timesheet)} resource+task+project timesheet combinations")
        logger.info(f"Found {len(designer_name_to_timesheet)} name+task+project timesheet combinations")
        logger.info(f"Found {len(designer_name_only_to_timesheet)} unique designer names with timesheets")
                
        # Step 5: Find resource IDs from employee IDs
        # This is needed because planning uses resource.resource while timesheet uses hr.employee
        employee_to_resource = {}
        
        # Try to map employees to resources using user_id as the link
        for resource_id, resource in ref_data.get('resources', {}).items():
            if resource.get('user_id') and isinstance(resource.get('user_id'), list) and len(resource.get('user_id')) > 0:
                user_id = resource.get('user_id')[0]
                # Store user_id -> resource_id mapping
                employee_to_resource[user_id] = resource_id
        
        # Step 6: Generate report data
        report_data = []
        
        # Dictionary to group tasks by designer
        designers = {}
        
        for slot in planning_slots:
            # Get resource info
            resource_id = None
            if 'resource_id' in slot and slot['resource_id'] and isinstance(slot['resource_id'], list):
                resource_id = slot['resource_id'][0]
                resource_name = slot['resource_id'][1] if len(slot['resource_id']) > 1 else "Unknown"
            else:
                resource_name = "Unknown"
            
            # Get task ID
            task_id = None
            if 'task_id' in slot and slot['task_id'] and isinstance(slot['task_id'], list):
                task_id = slot['task_id'][0]
                task_name = "Unknown"
                if task_id in ref_data.get('tasks', {}):
                    task_name = ref_data['tasks'][task_id].get('name', 'Unknown')
            else:
                task_name = "Unknown"
            
            # Get project ID
            project_id = None
            if 'project_id' in slot and slot['project_id'] and isinstance(slot['project_id'], list):
                project_id = slot['project_id'][0]
                project_name = "Unknown"
                if project_id in ref_data.get('projects', {}):
                    project_name = ref_data['projects'][project_id].get('name', 'Unknown')
            else:
                project_name = "Unknown"
            
            # Check if this resource/employee has logged time for this specific task/project using tiered approach
            has_timesheet = False
            hours_logged = 0.0
            
            # First check: exact match by resource_id + task_id + project_id
            key = (resource_id, task_id, project_id)
            if key in resource_task_to_timesheet:
                hours_logged = resource_task_to_timesheet[key]['hours']
                
                # Get the user_id associated with the resource (if available)
                resource_user_id = None
                if resource_id in ref_data.get('resources', {}) and ref_data['resources'][resource_id].get('user_id'):
                    resource_details = ref_data['resources'][resource_id]
                    if isinstance(resource_details['user_id'], list) and len(resource_details['user_id']) > 0:
                        resource_user_id = resource_details['user_id'][0]
                    elif isinstance(resource_details['user_id'], int):
                        resource_user_id = resource_details['user_id']
                
                # Only consider it a valid timesheet if hours logged are greater than 0
                # AND the entry was created by the designer (their user_id is in the user_ids set)
                user_ids = resource_task_to_timesheet[key]['user_ids']
                has_timesheet = (hours_logged > 0) and (resource_user_id in user_ids if resource_user_id else False)
            
            # Second check: try matching by name + task_id + project_id
            if not has_timesheet and resource_name != "Unknown":
                name_key = (resource_name, task_id, project_id)
                if name_key in designer_name_to_timesheet:
                    hours_logged = designer_name_to_timesheet[name_key]['hours']
                    has_timesheet = hours_logged > 0
            
            # Last resort: check if designer has ANY timesheet for the day
            if not has_timesheet and resource_name != "Unknown":
                if resource_name in designer_name_only_to_timesheet:
                    hours_logged = designer_name_only_to_timesheet[resource_name]['hours']
                    has_timesheet = hours_logged > 0
            
            # Get other slot info for display
            slot_name = slot.get('name', 'Unnamed Slot')
            
            # Convert boolean values to strings
            if isinstance(slot_name, bool):
                slot_name = str(slot_name)
            
            # Get shift status for display
            shift_status = slot.get('x_studio_shift_status', 'Unknown')
            
            # Get client success member (create_uid)
            client_success_name = "Unknown"
            if 'create_uid' in slot and slot['create_uid'] and isinstance(slot['create_uid'], list):
                create_uid = slot['create_uid'][0]
                if create_uid in ref_data.get('users', {}):
                    client_success_name = ref_data['users'][create_uid].get('name', 'Unknown')
            
            # Format start and end times for display
            start_datetime = slot.get('start_datetime', '')
            end_datetime = slot.get('end_datetime', '')
            
            start_time = "Unknown"
            end_time = "Unknown"
            
            if start_datetime and isinstance(start_datetime, str):
                try:
                    # Convert string to datetime
                    start_dt = datetime.strptime(start_datetime, "%Y-%m-%d %H:%M:%S")
                    start_time = start_dt.strftime("%H:%M")
                except:
                    start_time = start_datetime
            
            if end_datetime and isinstance(end_datetime, str):
                try:
                    # Convert string to datetime
                    end_dt = datetime.strptime(end_datetime, "%Y-%m-%d %H:%M:%S")
                    end_time = end_dt.strftime("%H:%M")
                except:
                    end_time = end_datetime
            
            # Get time allocation
            allocated_hours = slot.get('allocated_hours', 0.0)
            
            # Extract task date from slot data
            task_date = None
            if start_datetime and isinstance(start_datetime, str):
                try:
                    # Convert string to datetime
                    task_date = datetime.strptime(start_datetime, "%Y-%m-%d %H:%M:%S").date()
                except:
                    # If parsing fails, use the selected date
                    task_date = selected_date
            else:
                # Fallback if no valid start_datetime
                task_date = selected_date

            # Get sub task link
            sub_task_link = ""
            raw_sub_task_link = slot.get('x_studio_sub_task_link', False)
            logger.info(f"Original sub_task_link: {sub_task_link}, Type: {type(sub_task_link)}")

            # If we have a valid relation field, construct a proper Odoo URL
            if isinstance(raw_sub_task_link, list) and len(raw_sub_task_link) > 0:
                sub_task_id = raw_sub_task_link[0]  # Get the ID (first element)
                if sub_task_id:
                    # Construct Odoo task URL
                    base_url = st.session_state.odoo_url
                    sub_task_link = f"{base_url}/web#id={sub_task_id}&model=project.task&view_type=form"
                    logger.info(f"Constructed Odoo URL for task: {sub_task_link}")
            else:
                # Use the original handling for other cases
                sub_task_link = raw_sub_task_link

            # Ensure we have a valid URL format if it's not empty
            if sub_task_link and isinstance(sub_task_link, str):
                # Add the protocol if missing but looks like a URL
                if not (sub_task_link.startswith('http') or sub_task_link.startswith('/')):
                    if '.' in sub_task_link and not sub_task_link.startswith('www.'):
                        sub_task_link = 'https://' + sub_task_link
                        logger.info(f"Added https:// to link: {sub_task_link}")
            # Calculate days since task date for urgency
            reference_point = selected_date
            days_since_task = (reference_point - task_date).days
            
            # Only include if:
            # 1. We want all slots (debug mode) OR it has no timesheet
            # 2. We don't need to check against reference_date because we're already querying from that date
            if (st.session_state.debug_mode or not has_timesheet):
                task_data = {
                    'Date': task_date.strftime("%Y-%m-%d"),
                    'Designer': str(resource_name),
                    'Project': str(project_name),
                    'Client Success Member': str(client_success_name),
                    'Task': str(task_name),
                    # 'Slot Name': str(slot_name),
                    'Start Time': str(start_time),
                    'End Time': str(end_time),
                    'Allocated Hours': float(allocated_hours),
                    # 'Shift Status': str(shift_status),
                    'Days Overdue': int(days_since_task),
                    'Urgency': 'High' if days_since_task >= 2 else ('Medium' if days_since_task == 1 else 'Low'),
                    'Sub_Task_Link': str(sub_task_link)  # Add this line
                }
                
                report_data.append(task_data)
                
                # Also add to designers dictionary for notifications
                if resource_name not in designers:
                    designers[resource_name] = []
                designers[resource_name].append(task_data)
        
        # Convert to DataFrame
        if report_data:
            # Ensure all values are properly converted to appropriate types
            for item in report_data:
                # Convert any boolean values to strings
                for key, value in item.items():
                    if isinstance(value, bool):
                        item[key] = str(value)
                    # Handle other problematic types if needed
                    elif value is None:
                        item[key] = ""
            
            df = pd.DataFrame(report_data)
            
            # Count missing entries based on what's in report_data when not in debug mode
            missing_count = len(report_data) if not st.session_state.debug_mode else len([slot for slot in planning_slots if not (
                # First check: exact match by resource_id + task_id + project_id
                (slot.get('resource_id') and slot.get('task_id') and slot.get('project_id') and 
                (slot['resource_id'][0], slot['task_id'][0], slot['project_id'][0]) in resource_task_to_timesheet and
                resource_task_to_timesheet[(slot['resource_id'][0], slot['task_id'][0], slot['project_id'][0])]['hours'] > 0) or
                
                # Second check: try matching by name + task_id + project_id
                (slot.get('resource_id') and slot.get('task_id') and slot.get('project_id') and
                len(slot['resource_id']) > 1 and
                (slot['resource_id'][1], slot['task_id'][0], slot['project_id'][0]) in designer_name_to_timesheet and
                designer_name_to_timesheet[(slot['resource_id'][1], slot['task_id'][0], slot['project_id'][0])]['hours'] > 0) or
                
                # Last resort: check if designer has ANY timesheet for the day
                (slot.get('resource_id') and len(slot['resource_id']) > 1 and
                slot['resource_id'][1] in designer_name_only_to_timesheet and
                designer_name_only_to_timesheet[slot['resource_id'][1]]['hours'] > 0)
            )])
            
            # Send email report if requested
            if send_email and (missing_count > 0 or st.session_state.debug_mode):
                email_sent = send_email_report(df, selected_date, missing_count, len(timesheet_entries), shift_status_filter, reference_date)
                if email_sent:
                    st.success(f"Email report sent to {st.session_state.email_recipient}")
                else:
                    st.warning("Failed to send email report. Check email settings.")
            
            # Send individual emails to designers if enabled
            if ((send_designer_emails or st.session_state.designer_emails_enabled) and 
                st.session_state.smtp_server and 
                st.session_state.smtp_port and
                st.session_state.smtp_username and 
                st.session_state.smtp_password and
                st.session_state.designer_email_mapping and
                (missing_count > 0 or st.session_state.debug_mode)):
                
                try:
                    # SMTP settings
                    smtp_settings = {
                        "server": st.session_state.smtp_server,
                        "port": st.session_state.smtp_port,
                        "username": st.session_state.smtp_username,
                        "password": st.session_state.smtp_password
                    }
                    
                    email_success_count = 0
                    email_fail_count = 0
                    
                    # Send email to each designer with missing timesheets
                    for designer, tasks in designers.items():
                        # Check if we have an email mapping for this designer
                        if designer in st.session_state.designer_email_mapping:
                            designer_email = st.session_state.designer_email_mapping[designer]
                            
                            # Send the email
                            email_sent = send_designer_email(
                                designer,
                                designer_email,
                                selected_date,
                                tasks,
                                smtp_settings
                            )
                            
                            if email_sent:
                                email_success_count += 1
                            else:
                                email_fail_count += 1
                        else:
                            logger.info(f"No email mapping found for designer {designer}")
                    
                    # Show summary
                    if email_success_count > 0:
                        st.success(f"Sent emails to {email_success_count} designers")
                    if email_fail_count > 0:
                        st.warning(f"Failed to send emails to {email_fail_count} designers")
                            
                except Exception as e:
                    error_details = traceback.format_exc()
                    logger.error(f"Error sending designer emails: {e}\n{error_details}")
                    st.warning(f"Error sending designer emails: {e}")
            
            # Send Teams webhook notifications if enabled
            if st.session_state.webhooks_enabled and (missing_count > 0 or st.session_state.debug_mode):
                webhook_success_count = 0
                webhook_fail_count = 0
                
                try:
                    # Option 1: Test mode - send all notifications to test webhook
                    if st.session_state.test_webhook_url and not st.session_state.designer_webhook_mapping:
                        # Collect all tasks from all designers into one message 
                        all_tasks = []
                        for designer, tasks in designers.items():
                            # Add designer name to each task
                            for task in tasks:
                                task_copy = task.copy()
                                task_copy['Designer'] = designer
                                all_tasks.append(task_copy)
                        
                        # Send to test webhook
                        webhook_sent = send_teams_webhook_notification(
                            "All Designers (Test Mode)",
                            st.session_state.test_webhook_url,
                            all_tasks,
                            selected_date
                        )
                        
                        if webhook_sent:
                            webhook_success_count += 1
                            st.success(f"Sent combined Teams notification to your test webhook")
                        else:
                            webhook_fail_count += 1
                            st.warning("Failed to send Teams notification to test webhook")
                    
                    # Option 2: Production mode - send to individual designers
                    else:
                        for designer, tasks in designers.items():
                            # First check if designer has a webhook mapping
                            if designer in st.session_state.designer_webhook_mapping:
                                webhook_url = st.session_state.designer_webhook_mapping[designer]
                                
                                # Send the webhook notification
                                webhook_sent = send_teams_webhook_notification(
                                    designer,
                                    webhook_url,
                                    tasks,
                                    selected_date
                                )
                                
                                if webhook_sent:
                                    webhook_success_count += 1
                                else:
                                    webhook_fail_count += 1
                        
                        # Show summary if any webhooks were processed
                        if webhook_success_count + webhook_fail_count > 0:
                            if webhook_success_count > 0:
                                st.success(f"Sent Teams webhook notifications to {webhook_success_count} designers")
                            if webhook_fail_count > 0:
                                st.warning(f"Failed to send Teams webhook notifications to {webhook_fail_count} designers")
                
                except Exception as e:
                    error_details = traceback.format_exc()
                    logger.error(f"Error sending Teams webhook notifications: {e}\n{error_details}")
                    st.warning(f"Error sending Teams webhook notifications: {e}")
            # Add after the Teams webhook notification section
            # Send Teams direct messages if enabled
            if st.session_state.teams_direct_msg_enabled and (missing_count > 0 or st.session_state.debug_mode):
                try:
                    sent, success_count, fail_count = send_designer_teams_direct_messages(designers, selected_date)
                    
                    if sent:
                        if success_count > 0:
                            st.success(f"Sent Teams direct messages to {success_count} designers")
                        if fail_count > 0:
                            st.warning(f"Failed to send Teams direct messages to {fail_count} designers")
                except Exception as e:
                    error_details = traceback.format_exc()
                    logger.error(f"Error sending Teams direct messages: {e}\n{error_details}")
                    st.warning(f"Error sending Teams direct messages: {e}")
            # Add right after the Teams direct messages section in the generate_missing_timesheet_report function 
            # (around line 1155, after "st.warning(f"Failed to send Teams direct messages to {fail_count} designers")")
            # Send manager notifications if enabled
            if st.session_state.manager_emails_enabled and (missing_count > 0 or st.session_state.debug_mode):
                try:
                    sent, success_count, fail_count = send_manager_notifications(
                        designers, 
                        selected_date,
                        reference_date
                    )
                    
                    if sent:
                        if success_count > 0:
                            st.success(f"Sent notifications to {success_count} managers")
                        if fail_count > 0:
                            st.warning(f"Failed to send notifications to {fail_count} managers")
                except Exception as e:
                    error_details = traceback.format_exc()
                    logger.error(f"Error sending manager notifications: {e}\n{error_details}")
                    st.warning(f"Error sending manager notifications: {e}")
            
            return df, missing_count, len(timesheet_entries)
        else:
            # Return empty DataFrame with columns
            empty_df = pd.DataFrame(columns=[
                'Date', 'Designer', 'Project', 'Client Success Member', 'Task', 
                'Start Time', 'End Time', 'Allocated Hours', 'Days Overdue', 'Urgency'
            ])
            
            # Send email for empty report if requested
            if send_email:
                email_sent = send_email_report(empty_df, selected_date, 0, len(timesheet_entries), shift_status_filter, reference_date)
                if email_sent:
                    st.success(f"Email report sent to {st.session_state.email_recipient}")
                else:
                    st.warning("Failed to send email report. Check email settings.")
                
            return empty_df, 0, len(timesheet_entries)
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error generating report: {e}\n{error_details}")
        st.error(f"Error generating report: {e}")
        st.session_state.last_error = error_details
        return pd.DataFrame(), 0, len(timesheet_entries) if 'timesheet_entries' in locals() else 0

def send_teams_webhook_notification(
        designer_name: str,
        webhook_url: str,
        tasks: list,
        report_date: date
    ):
    """
    Post a short adaptive-card–compatible message to Teams.
    """
    try:
        max_days_overdue = max(t.get("Days Overdue", 0) for t in tasks)
        one_day  = (max_days_overdue == 1)
        two_plus = (max_days_overdue >= 2)

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

        # convert task list to bullet points
        bullet_lines = []
        for t in tasks:
            bullet_lines.append(
                f"- **{t.get('Task','?')}** "
                f"(Project: {t.get('Project','?')}, "
                f"Assigned on: {t.get('Date','?')}, "
                f"CS: {t.get('Client Success Member','?')})"
            )
        bullets = "\n".join(bullet_lines)

        card_text = (f"{emoji} **{title}**\n\n"
                     f"{intro}\n\n{bullets}\n\n"
                     "If something's blocking you, let us know – we're here "
                     "to help!")

        # basic card
        payload = {"text": card_text}

        resp = requests.post(webhook_url, json=payload, timeout=10)
        if resp.status_code == 200:
            logger.info("Teams webhook sent for %s", designer_name)
            return True
        else:
            logger.error("Teams webhook failed: %s %s",
                         resp.status_code, resp.text)
            return False

    except Exception as exc:
        logger.error("send_teams_webhook_notification failed: %s",
                     exc, exc_info=True)
        return False


def main():
    params = st.query_params
    if "headless" in params:
        st.session_state.headless_mode = True
        if "date" in params:
            date_param = params["date"][0]
            if date_param == "today":
                report_date = datetime.now().date()
            else:
                try:
                    report_date = datetime.strptime(date_param, "%Y-%m-%d").date()
                except ValueError:
                    report_date = datetime.now().date()
        else:
            report_date = datetime.now().date()
            
        # Set params based on URL parameters
        send_email = "email" in params and params["email"][0].lower() == "true"
        send_designer_emails = "designer_emails" in params and params["designer_emails"][0].lower() == "true"
        
        if "shift_status" in params:
            shift_status = params["shift_status"][0].lower()
            if shift_status == "all":
                shift_status_filter = None
            elif shift_status == "forecasted":
                shift_status_filter = "Forecasted"
            else:
                shift_status_filter = "Planned"
        else:
            shift_status_filter = "Planned"
        
        # Connect to Odoo and run the report
        uid, models = authenticate_odoo(
            st.session_state.odoo_url,
            st.session_state.odoo_db,
            st.session_state.odoo_username,
            st.session_state.odoo_password
        )
        
        if uid and models:
            st.session_state.odoo_uid = uid
            st.session_state.odoo_models = models
            
            # Generate report and send email
            df, missing_count, timesheet_count = generate_missing_timesheet_report(
                report_date, 
                shift_status_filter,
                send_email,
                send_designer_emails
            )
            
            st.write("Headless report generated successfully")
            st.write(f"Found {missing_count} missing entries out of {timesheet_count} timesheet entries")
            # Add an early return to skip the rest of the UI rendering
            return

    st.title("Missing Timesheet Reporter (Planning-Focused)")
    
    # Add a debug container that's initially hidden
    if 'debug_container' not in st.session_state:
        st.session_state.debug_container = None
    if 'debug_messages' not in st.session_state:
        st.session_state.debug_messages = []
    
    # Create a function to show debug info in the UI
    def debug_print(message, level="info"):
        """Add a debug message to be displayed in the UI"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        st.session_state.debug_messages.append({
            "time": timestamp,
            "message": message,
            "level": level
        })
        # Also log to regular logger
        if level == "error":
            logger.error(message)
        elif level == "warning":
            logger.warning(message)
        else:
            logger.info(message)
    
    # Create sidebar for options
    with st.sidebar:
        st.title("Options")
        
        # Debug mode toggle
        st.session_state.debug_mode = st.checkbox("Debug Mode", st.session_state.debug_mode, 
                                                 help="Show all planning slots, not just those missing timesheets")
        
        # Shift Status filter
        st.subheader("Shift Status Filter")
        shift_status_filter = st.radio(
            "Show slots with shift status:",
            ["All", "Planned (Confirmed)", "Forecasted (Unconfirmed)"],
            index=1,  # Default to "Planned (Confirmed)"
            help="Filter planning slots by their shift status"
        )

        # Update the shift status filter value based on selection
        if shift_status_filter == "All":
            st.session_state.shift_status_filter = None
        elif shift_status_filter == "Planned (Confirmed)":
            st.session_state.shift_status_filter = "Planned"
        elif shift_status_filter == "Forecasted (Unconfirmed)":
            st.session_state.shift_status_filter = "Forecasted"
        
        # Reference date setting for filtering historical tasks
        st.subheader("Reference Date")
        st.session_state.reference_date = st.date_input(
            "Reference date (ignore tasks before this date)",
            st.session_state.reference_date,
            help="Tasks before this date will be ignored to avoid alerting on historical data"
        )
        
        # Email notification settings
        with st.expander("Email Notifications", expanded=False):
            st.session_state.email_enabled = st.checkbox("Enable Email Reports", st.session_state.email_enabled,
                                                        help="Send email reports when generated")
            
            st.session_state.email_recipient = st.text_input("Email Recipient", 
                                                            st.session_state.email_recipient,
                                                            help="Email address to send reports to")
            
            st.session_state.smtp_server = st.text_input("SMTP Server", 
                                                        st.session_state.smtp_server,
                                                        help="SMTP server for sending emails")
            
            st.session_state.smtp_port = st.number_input("SMTP Port", 
                                                        min_value=1, 
                                                        max_value=65535, 
                                                        value=st.session_state.smtp_port,
                                                        help="SMTP port (usually 587 for TLS)")
            
            st.session_state.smtp_username = st.text_input("SMTP Username", 
                                                          st.session_state.smtp_username,
                                                          help="Username for SMTP authentication")
            
            st.session_state.smtp_password = st.text_input("SMTP Password", 
                                                          type="password",
                                                          value=st.session_state.smtp_password,
                                                          help="Password for SMTP authentication")
            
            st.markdown("""
            **Note:** For Gmail, you need to use an App Password instead of your regular password. 
            [Learn more](https://support.google.com/accounts/answer/185833)
            """)
            
            if st.button("Test Email"):
                try:
                    if not st.session_state.smtp_username or not st.session_state.smtp_password:
                        st.error("SMTP credentials are required")
                    else:
                        # Create test email
                        msg = MIMEMultipart()
                        msg['From'] = st.session_state.smtp_username
                        msg['To'] = st.session_state.email_recipient
                        msg['Subject'] = "Test Email from Missing Timesheet Reporter"
                        
                        body = """
                        <html>
                        <body>
                        <h2>Test Email</h2>
                        <p>This is a test email from the Missing Timesheet Reporter tool.</p>
                        <p>If you're receiving this, your email settings are configured correctly!</p>
                        </body>
                        </html>
                        """
                        
                        msg.attach(MIMEText(body, 'html'))
                        
                        # Send email
                        server = smtplib.SMTP(st.session_state.smtp_server, st.session_state.smtp_port)
                        server.starttls()
                        server.login(st.session_state.smtp_username, st.session_state.smtp_password)
                        server.send_message(msg)
                        server.quit()
                        
                        st.success(f"Test email sent to {st.session_state.email_recipient}")
                except Exception as e:
                    st.error(f"Failed to send test email: {e}")
                    if st.session_state.debug_mode:
                        st.code(traceback.format_exc())
        
        # Designer Email Notifications
        with st.sidebar.expander("Designer Email Notifications", expanded=False):
            st.session_state.designer_emails_enabled = st.checkbox(
                "Enable Designer Email Notifications", 
                st.session_state.designer_emails_enabled,
                help="Send individual emails to designers about their missing timesheet entries"
            )
            
            st.markdown("### Designer Email Mapping")
            st.markdown("Map designer names to their email addresses:")
            
            # Allow adding new designer email mappings
            col1, col2 = st.columns(2)
            with col1:
                new_designer = st.text_input("Designer Name", key="new_designer_name")
            with col2:
                new_email = st.text_input("Email Address", key="new_designer_email")
            
            if st.button("Add Designer"):
                if new_designer and new_email:
                    st.session_state.designer_email_mapping[new_designer] = new_email
                    st.success(f"Added mapping for {new_designer}")
                else:
                    st.error("Please enter both designer name and email")
            
            # Display current mappings and allow removal
            if st.session_state.designer_email_mapping:
                st.markdown("### Current Mappings")
                for idx, (designer, email) in enumerate(st.session_state.designer_email_mapping.items()):
                    col1, col2, col3 = st.columns([3, 3, 1])
                    with col1:
                        st.text(designer)
                    with col2:
                        st.text(email)
                    with col3:
                        if st.button("Remove", key=f"remove_{idx}"):
                            del st.session_state.designer_email_mapping[designer]
                            st.rerun()
            
        # Add after the "Designer Email Notifications" section (around line 770)
        with st.sidebar.expander("Manager Notifications", expanded=False):
            st.session_state.manager_emails_enabled = st.checkbox(
                "Enable Manager Notifications", 
                value=st.session_state.manager_emails_enabled,
                help="Send email notifications to managers about their team members' missing timesheets"
            )
            
            st.markdown("### Manager Email Settings")
            st.info("Manager notifications use the same email settings as the main email notifications. Make sure email notifications are configured.")
            
            # The file is already in the repository
            st.success("✅ Employee data loaded from uhd_data.xlsx")
            
            # Test button
            if st.button("Test Manager Notifications"):
                if not st.session_state.email_enabled:
                    st.error("Please enable email notifications first")
                elif not (st.session_state.smtp_server and 
                        st.session_state.smtp_port and
                        st.session_state.smtp_username and 
                        st.session_state.smtp_password):
                    st.error("Please configure email settings first")
                else:
                    # Create test data
                    test_manager_name = "Sanad Feras Khaleel Zaqtan"
                    test_manager_email = "sanad.zaqtan@prezlab.com"
                    
                    # Create test designer task data
                    test_designers = {
                        "Test Designer": [
                            {
                                "Project": "Test Project",
                                "Task": "Test Task",
                                "Start Time": "09:00",
                                "End Time": "17:00",
                                "Allocated Hours": 8.0,
                                "Date": datetime.now().date().strftime("%Y-%m-%d"),
                                "Days Overdue": 1,
                                "Client Success Member": "Test Manager"
                            }
                        ]
                    }
                    
                    with st.spinner("Sending test manager notifications..."):
                        # Send test notification directly
                        email_sent = send_manager_email(
                            test_manager_name,
                            test_manager_email,
                            test_designers,
                            datetime.now().date()
                        )
                        
                        if email_sent:
                            st.success(f"Test notification sent to {test_manager_name} ({test_manager_email})")
                        else:
                            st.error(f"Failed to send test notification to {test_manager_name}")
            # Add option to import mappings from CSV
            st.markdown("### Import Mappings")
            uploaded_file = st.file_uploader("Upload CSV with designer mappings", type="csv")
            if uploaded_file is not None:
                try:
                    df = pd.read_csv(uploaded_file)
                    if 'Designer' in df.columns and 'Email' in df.columns:
                        for _, row in df.iterrows():
                            designer = row['Designer']
                            email = row['Email']
                            if designer and email:
                                st.session_state.designer_email_mapping[designer] = email
                        st.success(f"Imported {len(df)} designer mappings")
                    else:
                        st.error("CSV must have 'Designer' and 'Email' columns")
                except Exception as e:
                    st.error(f"Error importing mappings: {e}")
            
            st.markdown("### Download Template")
            if st.button("Download CSV Template"):
                # Create template CSV
                template_df = pd.DataFrame({
                    'Designer': ['Designer Name 1', 'Designer Name 2'],
                    'Email': ['designer1@example.com', 'designer2@example.com']
                })
                csv = template_df.to_csv(index=False)
                
                # Provide download
                st.download_button(
                    label="Download Template",
                    data=csv,
                    file_name="designer_email_template.csv",
                    mime="text/csv"
                )
            
            # Add test button
            st.markdown("### Test Designer Email")
            test_designer = st.selectbox(
                "Select Designer to Test", 
                options=list(st.session_state.designer_email_mapping.keys()) if st.session_state.designer_email_mapping else ["No designers added"]
            )
            
            if st.button("Send Test Email"):
                if not st.session_state.designer_email_mapping:
                    st.error("Please add at least one designer email mapping")
                elif not (st.session_state.smtp_server and st.session_state.smtp_port and 
                        st.session_state.smtp_username and st.session_state.smtp_password):
                    st.error("Please configure email settings in the Email Notifications section")
                else:
                    # Get test designer
                    designer_email = st.session_state.designer_email_mapping.get(test_designer)
                    
                    # Create test task
                    test_task = [{
                        "Project": "Test Project",
                        "Task": "Test Task",
                        "Start Time": "09:00",
                        "End Time": "17:00",
                        "Allocated Hours": 8.0,
                        "Date": datetime.now().date().strftime("%Y-%m-%d")
                    }]
                    
                    # Send test email
                    smtp_settings = {
                        "server": st.session_state.smtp_server,
                        "port": st.session_state.smtp_port,
                        "username": st.session_state.smtp_username,
                        "password": st.session_state.smtp_password
                    }
                    
                    success = send_designer_email(
                        test_designer,
                        designer_email,
                        datetime.now().date(),
                        test_task,
                        smtp_settings
                    )
                    
                    if success:
                        st.success(f"Test email sent to {test_designer} ({designer_email})")
                    else:
                        st.error("Failed to send test email. Check email settings and try again.")        
        # Teams Webhooks Configuration
        with st.sidebar.expander("Teams Webhooks (No Admin Required)", expanded=False):
            st.session_state.webhooks_enabled = st.checkbox(
                "Enable Teams Webhooks", 
                value=st.session_state.webhooks_enabled,
                help="Send notifications via Teams channel webhooks (no admin consent required)"
            )
            
            st.markdown("### Test Webhook")
            st.markdown("Configure your own webhook for testing:")
            
            st.session_state.test_webhook_url = st.text_input(
                "Your Webhook URL",
                value=st.session_state.test_webhook_url,
                type="password",
                help="Webhook URL from your test private channel"
            )
            
            # Designer webhook mapping section
            st.markdown("### Designer Webhook Mapping")
            st.markdown("Map designer names to their channel webhook URLs:")
            
            # Allow adding new webhook mappings
            col1, col2 = st.columns(2)
            with col1:
                new_designer = st.text_input("Designer Name", key="new_webhook_designer")
            with col2:
                new_webhook = st.text_input("Webhook URL", key="new_webhook_url", type="password")
            
            if st.button("Add Webhook Mapping"):
                if new_designer and new_webhook:
                    st.session_state.designer_webhook_mapping[new_designer] = new_webhook
                    st.success(f"Added webhook mapping for {new_designer}")
                else:
                    st.error("Please enter both designer name and webhook URL")
            
            # Display current mappings and allow removal
            if st.session_state.designer_webhook_mapping:
                st.markdown("### Current Webhook Mappings")
                for idx, (designer, webhook) in enumerate(st.session_state.designer_webhook_mapping.items()):
                    col1, col2 = st.columns([4, 1])
                    with col1:
                        st.text(designer)
                    with col2:
                        if st.button("Remove", key=f"remove_webhook_{idx}"):
                            del st.session_state.designer_webhook_mapping[designer]
                            st.rerun()
            
            # Test button
            if st.button("Test Webhook"):
                if st.session_state.test_webhook_url:
                    # Create test task
                    test_task = [{
                        "Project": "Test Project",
                        "Task": "Test Task",
                        "Start Time": "09:00",
                        "End Time": "17:00",
                        "Allocated Hours": 8.0,
                        "Days Overdue": 2
                    }]
                    
                    # Send test
                    success = send_teams_webhook_notification(
                        "Test User",
                        st.session_state.test_webhook_url,
                        test_task,
                        datetime.now().date()
                    )
                    
                    if success:
                        st.success("Test message sent to your Teams channel!")
                    else:
                        st.error("Failed to send test message. Check the webhook URL.")
                else:
                    st.error("Please enter a webhook URL first")
        render_teams_direct_messaging_ui() 
        # Connection settings
        with st.expander("Connection Settings"):
            odoo_url = st.text_input("Odoo URL", st.session_state.odoo_url)
            odoo_db = st.text_input("Database", st.session_state.odoo_db)
            odoo_username = st.text_input("Username", st.session_state.odoo_username)
            odoo_password = st.text_input("Password", type="password", value=st.session_state.odoo_password)
            
            if st.button("Connect"):
                with st.spinner("Connecting to Odoo..."):
                    uid, models = authenticate_odoo(odoo_url, odoo_db, odoo_username, odoo_password)
                    
                    if uid and models:
                        st.session_state.odoo_uid = uid
                        st.session_state.odoo_models = models
                        st.session_state.odoo_db = odoo_db
                        st.session_state.odoo_url = odoo_url
                        st.session_state.odoo_username = odoo_username
                        st.session_state.odoo_password = odoo_password
                        st.success("Connected successfully!")
                    else:
                        st.error("Failed to connect to Odoo")
        # Add CSV data status expander right before the Connection Settings expander
        with st.sidebar.expander("CSV Data Status", expanded=False):
            if st.session_state.employee_data is not None:
                st.success(f"✅ Employee data loaded: {len(st.session_state.employee_data)} records")
                st.text(f"Email mappings: {len(st.session_state.designer_email_mapping)}")
                st.text(f"Teams ID mappings: {len(st.session_state.designer_teams_id_mapping)}")
                
                if st.button("Reload Data"):
                    st.session_state.employee_data = load_employee_data()
                    if st.session_state.employee_data is not None:
                        update_designer_mappings_from_csv()
                        st.success("Data reloaded successfully")
                    else:
                        st.error("Failed to reload data")
            else:
                st.error("❌ Employee data not loaded")
                if st.button("Try Loading Data"):
                    st.session_state.employee_data = load_employee_data()
                    if st.session_state.employee_data is not None:
                        update_designer_mappings_from_csv()
                        st.success("Data loaded successfully")
                    else:
                        st.error("Failed to load data")
    # Main area content
    # Connection status
    if st.session_state.odoo_uid and st.session_state.odoo_models:
        st.success(f"Connected to Odoo as {st.session_state.odoo_username}")
    else:
        st.warning("Not connected to Odoo. Please connect using the sidebar.")
        if st.button("Connect to Odoo"):
            with st.spinner("Connecting to Odoo..."):
                uid, models = authenticate_odoo(
                    st.session_state.odoo_url,
                    st.session_state.odoo_db,
                    st.session_state.odoo_username,
                    st.session_state.odoo_password
                )
                
                if uid and models:
                    st.session_state.odoo_uid = uid
                    st.session_state.odoo_models = models
                    st.success("Connected successfully!")
                else:
                    st.error("Failed to connect to Odoo")
    
    # Reference date info
    st.info(f"Date Range: {st.session_state.reference_date} to Selected Date - Showing unlogged hours in this range")
    
    # Date selection
    col1, col2, col3 = st.columns(3)
    
    with col1:
        selected_date = st.date_input(
            "Select End Date", 
            datetime.now().date(),  # Default to yesterday
            help="Choose the end date for the report range (reference date to this date)"
        )
    
    with col2:
        view_type = st.radio(
            "Group By", 
            ["Designer", "Project", "Urgency"],
            horizontal=True,
            help="Choose how to organize the report"
        )
        
    with col3:
        send_email_report = st.checkbox(
            "Send Email Report", 
            value=st.session_state.email_enabled,
            help="Send the report by email after generation"
        )
    
    # Generate report button
    if st.button("Generate Report", type="primary"):
        if not st.session_state.odoo_uid or not st.session_state.odoo_models:
            st.error("Not connected to Odoo. Please connect first.")
        else:
            # Display shift status filter info
            if st.session_state.shift_status_filter:
                shift_label = "Planned (Confirmed)" if st.session_state.shift_status_filter == "Planned" else "Forecasted (Unconfirmed)"
                st.info(f"Filtering for {shift_label} planning slots")
            
            with st.spinner("Generating timesheet report..."):
                # Generate the report
                df, missing_count, timesheet_count = generate_missing_timesheet_report(
                    selected_date, 
                    st.session_state.shift_status_filter,
                    send_email_report
                )
                
                # Display results
                if timesheet_count > 0:
                    st.info(f"Found {timesheet_count} timesheet entries from {st.session_state.reference_date.strftime('%Y-%m-%d')} to {selected_date.strftime('%Y-%m-%d')}")
                else:
                    st.warning(f"No timesheet entries found from {st.session_state.reference_date.strftime('%Y-%m-%d')} to {selected_date.strftime('%Y-%m-%d')}")
                    
                if missing_count > 0:
                    filter_text = ""
                    if st.session_state.shift_status_filter:
                        filter_text = f" with shift status '{st.session_state.shift_status_filter}'"
                    
                    st.warning(f"Found {missing_count} planning slots{filter_text} without timesheet entries from {st.session_state.reference_date.strftime('%Y-%m-%d')} to {selected_date.strftime('%Y-%m-%d')}")
                    
                    if not df.empty:
                        # Display data based on grouping
                        if view_type == "Designer":
                            st.subheader("Planning Slots by Designer")
                            
                            designer_summary = df.groupby("Designer").size().reset_index(name="Missing Entries")
                            designer_summary = designer_summary.sort_values("Missing Entries", ascending=False)
                            
                            st.dataframe(designer_summary)
                            
                            for designer in designer_summary["Designer"]:
                                designer_entries = df[df["Designer"] == designer]
                                with st.expander(f"{designer} - {len(designer_entries)} planning slots"):
                                    st.dataframe(designer_entries.drop(columns=["Designer"]))
                        
                        elif view_type == "Project":
                            st.subheader("Planning Slots by Project")
                            
                            project_summary = df.groupby("Project").size().reset_index(name="Missing Entries")
                            project_summary = project_summary.sort_values("Missing Entries", ascending=False)
                            
                            st.dataframe(project_summary)
                            
                            for project in project_summary["Project"]:
                                project_entries = df[df["Project"] == project]
                                with st.expander(f"{project} - {len(project_entries)} planning slots"):
                                    st.dataframe(project_entries.drop(columns=["Project"]))
                        
                        elif view_type == "Urgency":
                            st.subheader("Planning Slots by Urgency")
                            
                            # Apply color coding to urgency
                            def highlight_urgency(s):
                                if s == 'High':
                                    return 'background-color: red; color: white'
                                elif s == 'Medium':
                                    return 'background-color: orange; color: black'
                                else:
                                    return 'background-color: green; color: white'
                            
                            urgency_summary = df.groupby("Urgency").size().reset_index(name="Missing Entries")
                            urgency_order = {'High': 0, 'Medium': 1, 'Low': 2}
                            urgency_summary['order'] = urgency_summary['Urgency'].map(urgency_order)
                            urgency_summary = urgency_summary.sort_values("order")
                            urgency_summary = urgency_summary.drop(columns=["order"])
                            
                            st.dataframe(urgency_summary.style.applymap(highlight_urgency, subset=['Urgency']))
                            
                            # Display entries for each urgency level
                            for urgency in ['High', 'Medium', 'Low']:
                                if urgency in df['Urgency'].values:
                                    urgency_entries = df[df["Urgency"] == urgency]
                                    urgency_color = "red" if urgency == "High" else ("orange" if urgency == "Medium" else "green")
                                    
                                    with st.expander(f"{urgency} Urgency - {len(urgency_entries)} planning slots", expanded=(urgency == "High")):
                                        if urgency == "High":
                                            st.warning("These tasks are 2+ days overdue. Managers will be alerted if not addressed.")
                                        elif urgency == "Medium":
                                            st.info("These tasks are 1 day overdue and need immediate attention.")
                                            
                                        st.dataframe(urgency_entries)
                        
                        # Download button
                        st.subheader("Download Report")
                        
                        csv = df.to_csv(index=False)
                        st.download_button(
                            label="Download CSV",
                            data=csv,
                            file_name=f"planning_timesheet_report_{st.session_state.reference_date.strftime('%Y-%m-%d')}_to_{selected_date.strftime('%Y-%m-%d')}.csv",
                            mime="text/csv"
                        )
                else:
                    filter_text = ""
                    if st.session_state.shift_status_filter:
                        filter_text = f" (shift status: '{st.session_state.shift_status_filter}')"
                        
                    st.success(f"All planning slots{filter_text} have corresponding timesheet entries!")
                    
                    if not df.empty and st.session_state.debug_mode:
                        st.info("Showing all planning slots (debug mode)")
                        st.dataframe(df)
                    elif df.empty:
                        st.warning(f"No planning slots found for this date{filter_text}.")
            
            # Show error details if any
            if st.session_state.last_error and st.session_state.debug_mode:
                with st.expander("Error Details", expanded=False):
                    st.code(st.session_state.last_error)
    
    # Add a section for scheduled reports
    st.subheader("Schedule Daily Reports")
    st.info("""
    To set up scheduled daily reports, you need to run this app with a scheduled task or cron job.
    
    Example command for running a daily report at 9:00 AM:
    ```
    streamlit run app.py -- --headless --date=today --email=true
    ```
    
    Add this command to your server's cron jobs or task scheduler.
    """)
    # Automatic Scheduling Configuration
    st.subheader("Automatic Scheduling Configuration")

    # Container for scheduling settings
    schedule_container = st.container()
    with schedule_container:
        # Initialize session state variables for scheduling if they don't exist
        if 'auto_scheduling_enabled' not in st.session_state:
            st.session_state.auto_scheduling_enabled = False
        if 'auto_schedule_time' not in st.session_state:
            st.session_state.auto_schedule_time = "09:00"
        
        # UI elements for scheduling
        enable_col, time_col = st.columns([1, 2])
        
        with enable_col:
            auto_scheduling = st.checkbox("Enable Auto-Scheduling", 
                                        value=st.session_state.auto_scheduling_enabled,
                                        help="When enabled, the system will attempt to run reports automatically")
        
        with time_col:
            schedule_time = st.time_input("Daily Run Time", 
                                        value=datetime.strptime(st.session_state.auto_schedule_time, "%H:%M").time(),
                                        help="Time to run the report daily")
        
        # Save settings button
        if st.button("Save Scheduling Settings"):
            st.session_state.auto_scheduling_enabled = auto_scheduling
            st.session_state.auto_schedule_time = schedule_time.strftime("%H:%M")
            
            # Use a hardcoded app URL
            base_url = "https://prezlab-uhd.streamlit.app"
            schedule_url = f"{base_url}?headless=true&date=today&email=true&designer_emails=true&shift_status=planned"
            
            if auto_scheduling:
                st.success(f"Automatic scheduling enabled! Reports will run daily at {schedule_time.strftime('%H:%M')}.")
                st.info("Use this URL with an external scheduler service like cron-job.org:")
                st.code(schedule_url)
            else:
                st.warning("Automatic scheduling is disabled.")
if __name__ == "__main__":
    # Check if running in headless mode with command-line arguments
    import sys
    import argparse
    
    if len(sys.argv) > 1 and "--headless" in sys.argv:
        # Parse command-line arguments for headless mode
        parser = argparse.ArgumentParser(description="Missing Timesheet Reporter")
        parser.add_argument("--headless", action="store_true", help="Run in headless mode")
        parser.add_argument("--date", default="today", help="Date for report (YYYY-MM-DD or 'today')")
        parser.add_argument("--email", action="store_true", help="Send email report")
        parser.add_argument("--shift-status", default="planned", help="Shift status filter (planned, forecasted, or all)")
        parser.add_argument("--designer-emails", action="store_true", help="Send individual emails to designers")
        
        # Need to filter out Streamlit's own arguments
        streamlit_args = []
        our_args = []
        for arg in sys.argv[1:]:
            if arg.startswith("--headless") or arg.startswith("--date") or arg.startswith("--email") or arg.startswith("--shift-status"):
                our_args.append(arg)
            else:
                streamlit_args.append(arg)
                
        args, _ = parser.parse_known_args(our_args)
        
        # Set up date
        if args.date == "today":
            report_date = datetime.now().date()
        else:
            try:
                report_date = datetime.strptime(args.date, "%Y-%m-%d").date()
            except ValueError:
                logger.error(f"Invalid date format: {args.date}. Using today's date.")
                report_date = datetime.now().date()
        
        # Set shift status filter
        if args.shift_status.lower() == "all":
            shift_status_filter = None
        elif args.shift_status.lower() == "forecasted":
            shift_status_filter = "Forecasted"
        else:
            shift_status_filter = "Planned"  # Default
        
        logger.info(f"Running in headless mode for date {report_date}")
        
        # Connect to Odoo (using stored credentials)
        uid, models = authenticate_odoo(
            st.session_state.odoo_url,
            st.session_state.odoo_db,
            st.session_state.odoo_username,
            st.session_state.odoo_password
        )
        
        if uid and models:
            st.session_state.odoo_uid = uid
            st.session_state.odoo_models = models
            
            # Generate report and send email
            df, missing_count, timesheet_count = generate_missing_timesheet_report(
                report_date, 
                shift_status_filter,
                args.email  # Send email if --email flag is provided
            )
            
            logger.info(f"Report generated: {missing_count} missing entries out of {len(df)} total planning slots")
        else:
            logger.error("Failed to connect to Odoo in headless mode")
        
        # Exit after generating report in headless mode
        sys.exit(0)
    
    # Normal mode - run the Streamlit app
    main()