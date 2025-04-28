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

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='planning_timesheet_reporter.log'
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
if 'shift_status_filter' not in st.session_state:
    st.session_state.shift_status_filter = "confirmed"  # Default to Planned (confirmed)
if 'shift_status_values' not in st.session_state:
    st.session_state.shift_status_values = {}
if 'model_fields_cache' not in st.session_state:
    st.session_state.model_fields_cache = {}
if 'last_error' not in st.session_state:
    st.session_state.last_error = None
# Email settings
if 'email_enabled' not in st.session_state:
    st.session_state.email_enabled = False
if 'email_recipient' not in st.session_state:
    st.session_state.email_recipient = odoo_username  # Default to the Odoo username (usually an email)
if 'designer_emails_enabled' not in st.session_state:
    st.session_state.designer_emails_enabled = False
if 'designer_email_mapping' not in st.session_state:
    st.session_state.designer_email_mapping = {}
if 'smtp_server' not in st.session_state:
    st.session_state.smtp_server = "smtp.gmail.com"
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
        # Get shift status values if connected successfully
        if uid and models:
            shift_values = get_field_selection_values(
                models, uid, odoo_db, password, 'planning.slot', 'x_studio_shift_status'
            )
            st.session_state.shift_status_values = shift_values
            logger.info(f"Shift status values: {shift_values}")
        return uid, models
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Odoo connection error: {e}\n{error_details}")
        st.error(f"Odoo connection error: {e}")
        st.session_state.last_error = error_details
        return None, None
    
def get_field_selection_values(models, uid, odoo_db, odoo_password, model_name, field_name):
    """Get all possible selection values for a field"""
    try:
        # Get field definition
        fields_info = get_model_fields(models, uid, odoo_db, odoo_password, model_name)
        
        # Check if the field exists and is a selection type
        if field_name in fields_info and fields_info[field_name]['type'] == 'selection':
            # Return the selection options
            return dict(fields_info[field_name]['selection'])
        else:
            logger.warning(f"Field {field_name} is not a selection type or doesn't exist")
            return {}
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error getting selection values: {e}\n{error_details}")
        return {}
    
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

# Update to get_planning_slots function
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
        
        # Add shift status filter if provided
        # Add shift status filter if provided
        domains = []
        if shift_status_filter and 'x_studio_shift_status' in available_fields:
            logger.info(f"Filtering planning slots by x_studio_shift_status: {shift_status_filter}")
            
            # Generate multiple possible formats of the status value
            possible_values = [
                shift_status_filter,                      # Original value
                shift_status_filter.lower(),              # Lowercase
                shift_status_filter.upper(),              # Uppercase
                shift_status_filter.capitalize(),         # First letter capitalized
                shift_status_filter.title(),              # Title Case
            ]
            
            # Also try typical values if shift_status_filter contains certain keywords
            if "confirm" in shift_status_filter.lower():
                possible_values.extend(["Confirmed", "confirmed", "CONFIRMED"])
            if "plan" in shift_status_filter.lower():
                possible_values.extend(["Planned", "planned", "PLANNED"])
            if "forecast" in shift_status_filter.lower() or "unconfirm" in shift_status_filter.lower():
                possible_values.extend(["Forecasted", "forecasted", "FORECASTED", 
                                    "Unconfirmed", "unconfirmed", "UNCONFIRMED"])
            
            # Deduplicate
            possible_values = list(set(possible_values))
            
            # Log what we're trying
            logger.info(f"Trying these possible values for x_studio_shift_status: {possible_values}")
            
            # Create domains for each possible value
            for base_domain in base_domains:
                if base_domain:  # Skip empty domain
                    for value in possible_values:
                        domains.append(base_domain + [('x_studio_shift_status', '=', value)])

        else:
            domains = base_domains
        
        # Basic fields we want, checking which ones exist
        desired_fields = [
            'id', 'name', 'resource_id', 'start_datetime', 'end_datetime', 
            'allocated_hours', 'state', 'project_id', 'task_id', 'x_studio_shift_status',
            'create_uid', 'x_studio_sub_task_1', 'x_studio_task_activity', 'x_studio_service_category_1'
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
                        
                # Add shift status filter if provided
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
def debug_shift_status_values(models, uid, odoo_db, odoo_password):
    """Extract and display all possible values for x_studio_shift_status field with detailed logging"""
    try:
        st.subheader("Debugging Shift Status Values")
        
        # Get the field information
        fields_info = get_model_fields(models, uid, odoo_db, odoo_password, 'planning.slot')
        
        # Check if the field exists
        if 'x_studio_shift_status' not in fields_info:
            st.error("The field 'x_studio_shift_status' doesn't exist in the planning.slot model")
            logger.error("The field 'x_studio_shift_status' doesn't exist in the planning.slot model")
            return
        
        # Get field type and information
        field_info = fields_info['x_studio_shift_status']
        st.write(f"Field type: {field_info.get('type')}")
        
        # Show all field information for debugging
        st.json(field_info)
        
        # If it's a selection field, show all possible values
        if field_info.get('type') == 'selection':
            selection_values = field_info.get('selection', [])
            st.write("Possible values:")
            for value in selection_values:
                st.write(f"- Key: '{value[0]}', Label: '{value[1]}'")
            
            # Save to session state for reference
            st.session_state.shift_status_values = dict(selection_values)
        
        # Query actual values in the database
        st.subheader("Values currently in use")
        
        try:
            # Query distinct values in the database
            distinct_values = models.execute_kw(
                odoo_db, uid, odoo_password,
                'planning.slot', 'search_read',
                [[]],
                {'fields': ['x_studio_shift_status'], 'limit': 1000}
            )
            
            # Count occurrences of each value
            value_counts = {}
            for record in distinct_values:
                status = record.get('x_studio_shift_status', 'None')
                value_counts[status] = value_counts.get(status, 0) + 1
            
            # Display counts
            for status, count in sorted(value_counts.items(), key=lambda x: x[1], reverse=True):
                st.write(f"- '{status}': {count} records")
            
            # Also check for any raw database values
            st.subheader("Advanced: Raw Database Query")
            st.write("This section attempts a lower-level query to get all possible values")
            
            try:
                # Try a direct SQL query if the ORM approach isn't showing all values
                # Note: This may not work depending on Odoo's XML-RPC permissions
                query = """
                SELECT DISTINCT x_studio_shift_status, COUNT(*) as count
                FROM planning_slot
                GROUP BY x_studio_shift_status
                ORDER BY count DESC
                """
                
                sql_result = models.execute_kw(
                    odoo_db, uid, odoo_password,
                    'planning.slot', 'query', [query]
                )
                
                st.write("SQL Query Result:")
                st.json(sql_result)
            except Exception as e:
                st.warning(f"Direct SQL query not supported: {e}")
                
        except Exception as e:
            st.error(f"Error querying distinct values: {e}")
            logger.error(f"Error querying distinct values: {e}")
    
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error debugging shift status values: {e}\n{error_details}")
        st.error(f"Error debugging shift status values: {e}")
        st.session_state.last_error = error_details


def enhanced_shift_status_filtering(planning_slots, shift_status_filter, shift_status_values=None):
    """
    Enhanced filtering logic for planning slots based on shift status.
    This handles different possible representations and values for shift status.
    
    Args:
        planning_slots: List of planning slot dictionaries
        shift_status_filter: The status to filter by (e.g., "confirmed", "forecasted")
        shift_status_values: Dictionary mapping database values to display values (optional)
        
    Returns:
        List of filtered planning slots
    """
    if not shift_status_filter or not planning_slots:
        return planning_slots
    
    logger.info(f"Filtering {len(planning_slots)} slots with enhanced filter: '{shift_status_filter}'")
    
    # Convert shift_status_filter to lowercase for case-insensitive matching
    shift_lower = shift_status_filter.lower()
    
    # Track what values we've seen for debugging
    all_status_values = set()
    matched_values = set()
    
    # If we have shift_status_values from the database, use them for reverse lookup
    db_keys = []
    if shift_status_values:
        # Find all database keys that might correspond to our filter
        for db_key, label in shift_status_values.items():
            all_status_values.add(f"{db_key}:{label}")
            label_lower = label.lower()
            if shift_lower in label_lower or label_lower in shift_lower:
                db_keys.append(db_key)
                logger.info(f"Found matching database key: '{db_key}' for label '{label}'")
    
    # Filter the slots
    filtered_slots = []
    for slot in planning_slots:
        # Get the status value - could be a string or other format
        raw_status = slot.get('x_studio_shift_status')
        
        # Track all values we see
        all_status_values.add(str(raw_status))
        
        # Convert to string if it's not
        slot_status = str(raw_status).lower() if raw_status is not None else ""
        
        # Skip empty status if filtering is active
        if not slot_status and shift_status_filter:
            continue
        
        # Check various matching conditions
        is_match = False
        
        # Direct match by database key
        if db_keys and raw_status in db_keys:
            is_match = True
            matched_values.add(f"db_key:{raw_status}")
        
        # String matching heuristics for common keywords
        elif "confirm" in shift_lower and ("confirm" in slot_status or "plan" in slot_status):
            is_match = True
            matched_values.add(f"keyword_confirm:{slot_status}")
        elif "plan" in shift_lower and "plan" in slot_status:
            is_match = True
            matched_values.add(f"keyword_plan:{slot_status}")
        elif any(term in shift_lower for term in ["unconfirm", "forecast"]) and \
             any(term in slot_status for term in ["unconfirm", "forecast"]):
            is_match = True
            matched_values.add(f"keyword_forecast:{slot_status}")
        elif shift_lower == slot_status:
            is_match = True
            matched_values.add(f"exact:{slot_status}")
        
        if is_match:
            filtered_slots.append(slot)
    
    logger.info(f"Enhanced filtering: started with {len(planning_slots)} slots, returned {len(filtered_slots)}")
    logger.info(f"All status values seen: {sorted(all_status_values)}")
    logger.info(f"Matched values: {sorted(matched_values)}")
    
    return filtered_slots

def query_distinct_field_values(models, uid, odoo_db, odoo_password, model_name, field_name, limit=1000):
    """
    Query all distinct values for a specific field in the database.
    This is more reliable than relying on selection field definitions which might be out of sync.
    
    Args:
        models: Odoo models object
        uid: User ID
        odoo_db: Database name
        odoo_password: Password
        model_name: Model to query (e.g., 'planning.slot')
        field_name: Field to get distinct values for (e.g., 'x_studio_shift_status')
        limit: Maximum number of records to retrieve
        
    Returns:
        Dictionary mapping values to their counts in the database
    """
    try:
        # Search for all records with this field
        records = models.execute_kw(
            odoo_db, uid, odoo_password,
            model_name, 'search_read',
            [[]], 
            {'fields': [field_name], 'limit': limit}
        )
        
        # Count the occurrences of each value
        value_counts = {}
        for record in records:
            value = record.get(field_name, None)
            value_str = str(value) if value is not None else 'None'
            value_counts[value_str] = value_counts.get(value_str, 0) + 1
        
        # Also store the original values (not just strings)
        original_values = {}
        for record in records:
            value = record.get(field_name, None)
            value_str = str(value) if value is not None else 'None'
            original_values[value_str] = value
        
        logger.info(f"Found {len(value_counts)} distinct values for {field_name} in {model_name}")
        logger.info(f"Values: {value_counts}")
        
        return {
            'counts': value_counts,
            'values': original_values
        }
        
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error querying distinct values: {e}\n{error_details}")
        return {'counts': {}, 'values': {}}

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

def send_designer_email(designer_name, designer_email, report_date, tasks, smtp_settings):
    """Send email to designer with missing timesheet entries"""
    try:
        # Check if we have all required SMTP settings
        required_settings = ['server', 'port', 'username', 'password']
        if not all(key in smtp_settings for key in required_settings):
            logger.error("Missing required SMTP settings")
            return False
            
        # Create email
        msg = MIMEMultipart()
        msg['From'] = smtp_settings['username']
        msg['To'] = designer_email
        
        # Find the most overdue task to set email urgency
        max_days_overdue = 0
        for task in tasks:
            if 'Days Overdue' in task:
                max_days_overdue = max(max_days_overdue, task['Days Overdue'])
            elif 'Date' in task:
                task_date = datetime.strptime(task['Date'], "%Y-%m-%d").date()
                days_overdue = (report_date - task_date).days
                max_days_overdue = max(max_days_overdue, days_overdue)
        
        # Set subject line based on days overdue
        if max_days_overdue >= 2:
            msg['Subject'] = "Heads-Up: You've Missed Logging Hours for 2 Days"
        else:
            msg['Subject'] = "Quick Nudge – Log Your Hours"
        
        # Create email body based on days overdue
        date_str = report_date.strftime('%Y-%m-%d')
        
        if max_days_overdue >= 2:
            # 2+ days overdue template
            body = f"""
            <html>
            <body style="font-family: Arial, sans-serif; line-height: 1.5;">
            <p>Hi {designer_name},</p>
            
            <p>It looks like no hours have been logged for the past two days for the following task(s):</p>
            """
            
            # Add task details in bullet points
            for task in tasks:
                # Get client success member
                client_success = task.get('Client Success Member', 'Not specified')
                
                # Get task dates (could be multiple dates)
                task_date = task.get('Date', date_str)
                
                body += f"""
                <ul style="list-style-type: disc; padding-left: 20px;">
                    <li><strong>Task:</strong> {task.get('Task', 'Unknown')}</li>
                    <li><strong>Project:</strong> {task.get('Project', 'Unknown')}</li>
                    <li><strong>Assignment Dates:</strong> {task_date}</li>
                    <li><strong>Client Success Contact:</strong> {client_success}</li>
                </ul>
                """
            
            body += """
            <p>We completely understand things can get busy — but consistent time logging helps us improve project planning and smooth reporting.</p>
            
            <p>If something's holding you back from logging your hours, just reach out. We're here to help.</p>
            </body>
            </html>
            """
            
        else:
            # 1 day overdue template
            body = f"""
            <html>
            <body style="font-family: Arial, sans-serif; line-height: 1.5;">
            <p>Hi {designer_name},</p>
            
            <p>This is a gentle reminder to log your hours for the task below — It takes a minute, but the impact is big:</p>
            """
            
            # Add task details in bullet points (for each task)
            for task in tasks:
                # Get client success member
                client_success = task.get('Client Success Member', 'Not specified')
                
                # Get task date and time
                task_date = task.get('Date', date_str)
                task_time = f"{task.get('Start Time', 'Unknown')} - {task.get('End Time', 'Unknown')}"
                
                body += f"""
                <ul style="list-style-type: disc; padding-left: 20px;">
                    <li><strong>Task:</strong> {task.get('Task', 'Unknown')}</li>
                    <li><strong>Project:</strong> {task.get('Project', 'Unknown')}</li>
                    <li><strong>Assigned on:</strong> {task_date} {task_time}</li>
                    <li><strong>Client Success Contact:</strong> {client_success}</li>
                </ul>
                """
            
            body += """
            <p>Taking a minute now helps us stay on top of things later 🙌</p>
            <p>Let us know if you need any support with this.</p>
            </body>
            </html>
            """
        
        # Attach email body
        msg.attach(MIMEText(body, 'html'))
        
        # Send email
        server = smtplib.SMTP(smtp_settings['server'], smtp_settings['port'])
        server.starttls()
        server.login(smtp_settings['username'], smtp_settings['password'])
        server.send_message(msg)
        server.quit()
        
        logger.info(f"Email alert sent to {designer_name} ({designer_email})")
        return True
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error sending email to {designer_name}: {e}\n{error_details}")
        return False

def generate_missing_timesheet_report(selected_date, shift_status_filter=None, send_email=False, send_designer_emails=False):
    """
    Generate report of planning slots without timesheet entries for a date range from reference_date to selected_date
    """
    uid = st.session_state.odoo_uid
    models = st.session_state.odoo_models
    odoo_db = st.session_state.odoo_db
    odoo_password = st.session_state.odoo_password
    reference_date = st.session_state.reference_date
    
    if uid and models:
        st.session_state.shift_status_values = get_field_selection_values(
            models, uid, odoo_db, odoo_password, 'planning.slot', 'x_studio_shift_status'
        )
        logger.info(f"Shift status values: {st.session_state.shift_status_values}")
    
    if not uid or not models:
        st.error("Not connected to Odoo")
        return pd.DataFrame(), 0, 0
    
    try:
        # Step 1: Get all planning slots for the date range (reference_date to selected_date) with optional allocation filter
        planning_slots = get_planning_slots(models, uid, odoo_db, odoo_password, reference_date, selected_date, shift_status_filter)
        
        # Apply enhanced filtering for shift status
        if shift_status_filter:
            planning_slots = enhanced_shift_status_filtering(
                planning_slots, 
                shift_status_filter,
                st.session_state.shift_status_values
            )
            logger.info(f"Enhanced filtering returned {len(planning_slots)} slots with status matching '{shift_status_filter}'")
                
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
            entry_date = entry.get('date', None)
            
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
            
            # Extract task date from slot data
            task_date = None
            if slot.get('start_datetime') and isinstance(slot.get('start_datetime'), str):
                try:
                    # Convert string to datetime
                    task_date = datetime.strptime(slot.get('start_datetime'), "%Y-%m-%d %H:%M:%S").date()
                except:
                    # If parsing fails, use the selected date
                    task_date = selected_date
            else:
                # Fallback if no valid start_datetime
                task_date = selected_date
            
            # Format the task date as a string for comparison with timesheet dates
            task_date_str = task_date.strftime("%Y-%m-%d")
            
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
                
                # Check if any of the timesheet entries match the specific date
                matching_entries = [
                    entry for entry in resource_task_to_timesheet[key]['entries']
                    if entry.get('date') == task_date_str
                ]
                
                if matching_entries:
                    # Calculate hours for just the matching date
                    date_specific_hours = sum(entry.get('unit_amount', 0) for entry in matching_entries)
                    has_timesheet = date_specific_hours > 0 
            
            # Second check: try matching by name + task_id + project_id
            if not has_timesheet and resource_name != "Unknown":
                name_key = (resource_name, task_id, project_id)
                if name_key in designer_name_to_timesheet:
                    # Check if any of the timesheet entries match the specific date
                    matching_entries = [
                        entry for entry in designer_name_to_timesheet[name_key]['entries']
                        if entry.get('date') == task_date_str
                    ]
                    
                    if matching_entries:
                        # Calculate hours for just the matching date
                        date_specific_hours = sum(entry.get('unit_amount', 0) for entry in matching_entries)
                        has_timesheet = date_specific_hours > 0
            
            # Last resort: check if designer has ANY timesheet for THIS SPECIFIC DAY
            # if not has_timesheet and resource_name != "Unknown" and resource_name in designer_name_only_to_timesheet:
            #     # Filter entries to only include those for the specific task date
            #     matching_entries = [
            #         entry for entry in designer_name_only_to_timesheet[resource_name]['entries']
            #         if entry.get('date') == task_date_str
            #     ]
                
            #     if matching_entries:
            #         # Calculate hours for just the matching date
            #         date_specific_hours = sum(entry.get('unit_amount', 0) for entry in matching_entries)
            #         has_timesheet = date_specific_hours > 0
            
            # Get other slot info for display
            slot_name = slot.get('name', 'Unnamed Slot')
            
            # Convert boolean values to strings
            if isinstance(slot_name, bool):
                slot_name = str(slot_name)
            
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
                    'Slot Name': str(slot_name),
                    'Start Time': str(start_time),
                    'End Time': str(end_time),
                    'Allocated Hours': float(allocated_hours),
                    'Shift Status': str(slot.get('x_studio_shift_status', 'Unknown')),  # Use actual status from slot
                    'Days Overdue': int(days_since_task),
                    'Urgency': 'High' if days_since_task >= 2 else ('Medium' if days_since_task == 1 else 'Low')
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
            missing_count = len(report_data)
            
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
            
            return df, missing_count, len(timesheet_entries)
        else:
            # Return empty DataFrame with columns
            empty_df = pd.DataFrame(columns=[
                'Date', 'Designer', 'Project', 'Client Success Member', 'Task', 'Slot Name', 
                'Start Time', 'End Time', 'Allocated Hours', 'Shift Status', 'Days Overdue', 'Urgency'
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

def send_teams_webhook_notification(designer_name, webhook_url, tasks, report_date):
    """Send a simplified notification to a Teams channel via webhook"""
    try:
        # Find the most overdue task
        max_days_overdue = 0
        for task in tasks:
            days_overdue = task.get('Days Overdue', 0)
            max_days_overdue = max(max_days_overdue, days_overdue)
        
        # Create date string
        date_str = report_date.strftime('%Y-%m-%d')
        
        # Example task for message (just use the first one)
        task = tasks[0] if tasks else {}
        task_name = task.get('Task', 'Unknown')
        project_name = task.get('Project', 'Unknown')
        client_success = task.get('Client Success Member', 'Not specified')
        task_date = task.get('Date', date_str)
        
        # Create message based on overdue status
        if max_days_overdue >= 2:
            message = {
                "text": f"**Heads-Up: You've Missed Logging Hours for 2 Days**\n\n"
                       f"Hi {designer_name},\n\n"
                       f"It looks like no hours have been logged for the past two days for the following task:\n\n"
                       f"* **Task:** {task_name}\n"
                       f"* **Project:** {project_name}\n"
                       f"* **Assignment Dates:** {task_date}\n"
                       f"* **Client Success Contact:** {client_success}\n\n"
                       f"We completely understand things can get busy — but consistent time logging helps us improve project planning and smooth reporting.\n\n"
                       f"If something's holding you back from logging your hours, just reach out. We're here to help."
            }
        else:
            # Format the time if available
            time_str = ""
            if 'Start Time' in task and 'End Time' in task:
                time_str = f"{task.get('Start Time', '')} - {task.get('End Time', '')}"
            
            message = {
                "text": f"**Quick Nudge – Log Your Hours**\n\n"
                       f"Hi {designer_name},\n\n"
                       f"This is a gentle reminder to log your hours for the task below — It takes a minute, but the impact is big:\n\n"
                       f"* **Task:** {task_name}\n"
                       f"* **Project:** {project_name}\n"
                       f"* **Assigned on:** {task_date} {time_str}\n"
                       f"* **Client Success Contact:** {client_success}\n\n"
                       f"Taking a minute now helps us stay on top of things later 🙌\n\n"
                       f"Let us know if you need any support with this."
            }
        
        # If there are multiple tasks, add a note
        if len(tasks) > 1:
            message["text"] += f"\n\n**Note:** There are {len(tasks)} tasks in total requiring attention."
        
        # Send to Teams webhook
        response = requests.post(webhook_url, json=message)
        
        if response.status_code == 200:
            logger.info(f"Teams webhook notification sent for {designer_name}")
            return True
        else:
            logger.error(f"Failed to send Teams webhook: {response.status_code} - {response.text}")
            return False
            
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error sending Teams webhook for {designer_name}: {e}\n{error_details}")
        return False

def main():
    st.title("Missing Timesheet Reporter (Planning-Focused)")
    
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

        # Map the human label → actual value stored in x_studio_shift_status
        if st.session_state.shift_status_values:
            # If we have values from the database, try to match them
            possible_values = {value.lower(): key for key, value in st.session_state.shift_status_values.items()}
            planned_value = None
            forecast_value = None
            for db_value, db_key in possible_values.items():
                if "confirm" in db_value or "plan" in db_value:
                    planned_value = db_key
                if "unconfirm" in db_value or "forecast" in db_value:
                    forecast_value = db_key
            
            # Use found values or fallbacks
            STATUS_MAP = {
                "Planned (Confirmed)": planned_value or "Planned",
                "Forecasted (Unconfirmed)": forecast_value or "Forecasted"
            }
        else:
            # Fallback to default values if we haven't loaded the values yet
            STATUS_MAP = {
                "Planned (Confirmed)": "confirmed",
                "Forecasted (Unconfirmed)": "unconfirmed"
            }

        # Log what values we're using
        logger.info(f"Using status map: {STATUS_MAP}")
        if shift_status_filter == "All":
            st.session_state.shift_status_filter = None
        else:
            st.session_state.shift_status_filter = STATUS_MAP[shift_status_filter]

        
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
                            st.experimental_rerun()
            
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
                            st.experimental_rerun()
            
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
                        
                        # Add this code here
                        shift_values = get_field_selection_values(
                            models, uid, odoo_db, odoo_password, 
                            'planning.slot', 'x_studio_shift_status'
                        )
                        st.session_state.shift_status_values = shift_values
                        logger.info(f"Shift status values: {shift_values}")
                    else:
                        st.error("Failed to connect to Odoo")
    
    # Main area content
    # Connection status
    if st.session_state.odoo_uid and st.session_state.odoo_models:
        st.success(f"Connected to Odoo as {st.session_state.odoo_username}")
        
        # Add the debug button right here, after the connection success message
        if st.session_state.debug_mode:
            if st.button("Show Shift Status Values"):
                if st.session_state.shift_status_values:
                    st.write("Available shift status values:", st.session_state.shift_status_values)
                    # Also show what would be mapped for UI labels
                    st.write("UI mapping:")
                    for label, value in {
                        "Planned (Confirmed)": "confirmed",
                        "Forecasted (Unconfirmed)": "unconfirmed"
                    }.items():
                        matching_db_value = None
                        for db_key, db_val in st.session_state.shift_status_values.items():
                            if value.lower() in db_val.lower() or db_val.lower() in value.lower():
                                matching_db_value = f"{db_key} ({db_val})"
                                break
                        st.write(f"  {label} → {matching_db_value or 'No match found'}")
                else:
                    st.write("No shift status values loaded yet")
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
                    
                    # Add this code here
                    shift_values = get_field_selection_values(
                        models, uid, st.session_state.odoo_db, st.session_state.odoo_password, 
                        'planning.slot', 'x_studio_shift_status'
                    )
                    st.session_state.shift_status_values = shift_values
                    logger.info(f"Shift status values: {shift_values}")
                else:
                    st.error("Failed to connect to Odoo")
    
    # Reference date info
    st.info(f"Date Range: {st.session_state.reference_date} to Selected Date - Showing unlogged hours in this range")
    
    # Date selection
    col1, col2, col3 = st.columns(3)
    
    with col1:
        selected_date = st.date_input(
            "Select End Date", 
            datetime.now().date() - timedelta(days=1),  # Default to yesterday
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
                shift_status_label = (
                    "Planned (Confirmed)" if st.session_state.shift_status_filter == "confirmed"
                    else "Forecasted (Unconfirmed)"
                )
                st.info(f"Filtering for {shift_status_label} planning slots")
            
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
                            file_name=f"planning_timesheet_report_{reference_date.strftime('%Y-%m-%d')}_to_{selected_date.strftime('%Y-%m-%d')}.csv",
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
    if st.session_state.debug_mode:
        st.markdown("---")
        st.subheader("🔍 Debugging Tools")
        
        debug_tab1, debug_tab2, debug_tab3 = st.tabs(["Shift Status Values", "Planning Slots Query", "Timesheet Matching"])
        
        with debug_tab1:
            st.markdown("### Debug Shift Status Values")
            st.write("This will extract all possible values for the x_studio_shift_status field")
            
            if st.button("Extract Shift Status Values"):
                if not st.session_state.odoo_uid or not st.session_state.odoo_models:
                    st.error("Not connected to Odoo. Please connect first.")
                else:
                    with st.spinner("Extracting shift status values..."):
                        debug_shift_status_values(
                            st.session_state.odoo_models,
                            st.session_state.odoo_uid,
                            st.session_state.odoo_db,
                            st.session_state.odoo_password
                        )
            
            # Display current mapping
            if st.session_state.shift_status_values:
                st.write("Current shift status values in memory:")
                for key, value in st.session_state.shift_status_values.items():
                    st.write(f"- Key: '{key}', Label: '{value}'")
            
            # Option to manually set mapping
            st.markdown("### Manual Mapping Override")
            st.write("If automatic mapping isn't working, you can manually define the mappings:")
            
            col1, col2 = st.columns(2)
            with col1:
                confirmed_key = st.text_input("Key for 'Confirmed' status:", 
                                            key="confirmed_key_override")
            with col2:
                forecasted_key = st.text_input("Key for 'Forecasted' status:", 
                                            key="forecasted_key_override")
            
            if st.button("Apply Manual Mapping"):
                if confirmed_key or forecasted_key:
                    # Create new mapping dictionary
                    manual_mapping = {}
                    if confirmed_key:
                        manual_mapping[confirmed_key] = "Confirmed"
                    if forecasted_key:
                        manual_mapping[forecasted_key] = "Forecasted"
                    
                    # Update session state
                    st.session_state.shift_status_values = manual_mapping
                    st.success("Manual mapping applied!")
                    
            # Query actual database values
            st.markdown("### Query Actual Database Values")
            if st.button("Query Distinct Shift Status Values"):
                if not st.session_state.odoo_uid or not st.session_state.odoo_models:
                    st.error("Not connected to Odoo. Please connect first.")
                else:
                    with st.spinner("Querying database..."):
                        result = query_distinct_field_values(
                            st.session_state.odoo_models,
                            st.session_state.odoo_uid,
                            st.session_state.odoo_db,
                            st.session_state.odoo_password,
                            'planning.slot',
                            'x_studio_shift_status',
                            5000  # Increased limit to get more comprehensive results
                        )
                        
                        if result['counts']:
                            st.write("Values found in the database:")
                            for value, count in sorted(result['counts'].items(), key=lambda x: x[1], reverse=True):
                                st.write(f"- '{value}': {count} records")
                            
                            # Automatically create mappings based on keywords
                            st.write("Suggested mappings:")
                            confirmed_key = None
                            forecasted_key = None
                            
                            for value_str in result['counts'].keys():
                                value_lower = value_str.lower()
                                if "confirm" in value_lower or "plan" in value_lower:
                                    confirmed_key = result['values'][value_str]
                                    st.write(f"- Confirmed/Planned: '{value_str}'")
                                elif "unconfirm" in value_lower or "forecast" in value_lower:
                                    forecasted_key = result['values'][value_str]
                                    st.write(f"- Forecasted/Unconfirmed: '{value_str}'")
                            
                            # Option to apply suggested mappings
                            if confirmed_key is not None or forecasted_key is not None:
                                if st.button("Apply Suggested Mappings"):
                                    mapping = {}
                                    if confirmed_key is not None:
                                        mapping[confirmed_key] = "Confirmed"
                                    if forecasted_key is not None:
                                        mapping[forecasted_key] = "Forecasted"
                                        
                                    st.session_state.shift_status_values = mapping
                                    st.success("Applied suggested mappings!")
                        else:
                            st.warning("No values found or error occurred")
        
        with debug_tab2:
            st.markdown("### Debug Planning Slot Query")
            st.write("Test direct query for planning slots with specific settings")
            
            # Date range for test
            test_start_date = st.date_input("Test Start Date", 
                                        value=datetime.now().date() - timedelta(days=7),
                                        key="test_start_date")
            
            test_end_date = st.date_input("Test End Date", 
                                        value=datetime.now().date(),
                                        key="test_end_date")
            
            # Status filter options
            test_status = st.selectbox("Test Status Filter",
                                    options=["None"] + list(st.session_state.shift_status_values.values()) 
                                    if st.session_state.shift_status_values 
                                    else ["None", "confirmed", "forecasted", "planned", "unconfirmed"],
                                    key="test_status")
            
            if test_status == "None":
                test_status = None
            
            if st.button("Test Planning Slot Query"):
                if not st.session_state.odoo_uid or not st.session_state.odoo_models:
                    st.error("Not connected to Odoo. Please connect first.")
                else:
                    with st.spinner("Querying planning slots..."):
                        test_slots = get_planning_slots(
                            st.session_state.odoo_models,
                            st.session_state.odoo_uid,
                            st.session_state.odoo_db,
                            st.session_state.odoo_password,
                            test_start_date,
                            test_end_date,
                            test_status
                        )
                        
                        st.write(f"Found {len(test_slots)} planning slots")
                        
                        # Show sample of the slots
                        if test_slots:
                            sample_size = min(5, len(test_slots))
                            st.write(f"Sample of {sample_size} slots:")
                            for i in range(sample_size):
                                st.json(test_slots[i])
                            
                            # Count by status
                            status_counts = {}
                            for slot in test_slots:
                                status = slot.get('x_studio_shift_status', 'None')
                                status_counts[status] = status_counts.get(status, 0) + 1
                            
                            st.write("Status distribution:")
                            for status, count in status_counts.items():
                                st.write(f"- '{status}': {count} slots")
        
        with debug_tab3:
            st.markdown("### Debug Timesheet Matching")
            st.write("Test the logic that matches timesheets to planning slots")
            
            # Simple test with a specific date
            test_match_date = st.date_input("Test Date for Matching", 
                                            value=datetime.now().date() - timedelta(days=1),
                                            key="test_match_date")
            
            if st.button("Test Timesheet Matching"):
                if not st.session_state.odoo_uid or not st.session_state.odoo_models:
                    st.error("Not connected to Odoo. Please connect first.")
                else:
                    with st.spinner("Testing timesheet matching..."):
                        # Get planning slots
                        test_slots = get_planning_slots(
                            st.session_state.odoo_models,
                            st.session_state.odoo_uid,
                            st.session_state.odoo_db,
                            st.session_state.odoo_password,
                            test_match_date,
                            test_match_date,
                            None  # No status filter
                        )
                        
                        # Get timesheet entries
                        test_entries = get_timesheet_entries(
                            st.session_state.odoo_models,
                            st.session_state.odoo_uid,
                            st.session_state.odoo_db,
                            st.session_state.odoo_password,
                            test_match_date,
                            test_match_date
                        )
                        
                        st.write(f"Found {len(test_slots)} planning slots and {len(test_entries)} timesheet entries")
                        
                        # Display some examples of each
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.markdown("### Sample Planning Slots")
                            if test_slots:
                                for i in range(min(3, len(test_slots))):
                                    with st.expander(f"Slot {i+1}"):
                                        st.json(test_slots[i])
                            else:
                                st.write("No planning slots found")
                        
                        with col2:
                            st.markdown("### Sample Timesheet Entries")
                            if test_entries:
                                for i in range(min(3, len(test_entries))):
                                    with st.expander(f"Entry {i+1}"):
                                        st.json(test_entries[i])
                            else:
                                st.write("No timesheet entries found")
                        
                        # Show raw field formats for comparison
                        st.markdown("### Field Formats Comparison")
                        
                        formats = {
                            "Planning Slot": {
                                "resource_id": test_slots[0].get("resource_id") if test_slots else None,
                                "task_id": test_slots[0].get("task_id") if test_slots else None,
                                "project_id": test_slots[0].get("project_id") if test_slots else None,
                                "x_studio_shift_status": test_slots[0].get("x_studio_shift_status") if test_slots else None
                            },
                            "Timesheet Entry": {
                                "employee_id": test_entries[0].get("employee_id") if test_entries else None,
                                "task_id": test_entries[0].get("task_id") if test_entries else None,
                                "project_id": test_entries[0].get("project_id") if test_entries else None,
                                "user_id": test_entries[0].get("user_id") if test_entries else None
                            }
                        }
                        
                        st.json(formats)
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
        parser.add_argument("--shift-status", default="forecasted", help="Shift status filter (forecasted, planned, or all)")
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
            shift_status_filter = "forecasted"
        else:
            shift_status_filter = "planned"  # Default
        
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
                st.session_state.shift_status_filter,
                args.email  # Send email if --email flag is provided
            )
            
            logger.info(f"Report generated: {missing_count} missing entries out of {len(df)} total planning slots")
        else:
            logger.error("Failed to connect to Odoo in headless mode")
        
        # Exit after generating report in headless mode
        sys.exit(0)
    
    # Normal mode - run the Streamlit app
    main()