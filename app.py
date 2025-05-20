import os
import sys

# ─── Make sure local modules and the data folder are importable ─────────────────
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# Now try to import
try:
    from config import get_secret
except ImportError as e:
    print(f"Error importing config: {e}")
    # Fallback implementation
    def get_secret(key, default=None):
        """Fallback get_secret implementation"""
        if hasattr(st, 'secrets') and key in st.secrets:
            return st.secrets[key]
        return os.environ.get(key, default)
    
# ─── Standard library ─────────────────────────────────────────────────────────
import logging
import traceback
import re
import uuid
from pathlib import Path
from datetime import datetime, date, time
from typing import Any, Dict, List, Optional, Tuple, Union

# ─── Third‑party ──────────────────────────────────────────────────────────────
import streamlit as st
import pandas as pd

# ─── Local modules ───────────────────────────────────────────────────────────
from config import get_secret

# Debug util (graceful fallback if missing)
try:
    from debug_utils import inject_debug_page, debug_function, SystemDebugger
except ImportError:
    def inject_debug_page(): return False
    def debug_function(f): return f
    class SystemDebugger:
        def streamlit_debug_page(self): pass

# Core helpers (all secrets now loaded inside these functions)
from helpers import (
    authenticate_odoo,
    create_odoo_task,
    get_sales_orders,
    get_sales_order_details,
    get_employee_schedule,
    create_task,
    find_employee_id,
    get_target_languages_odoo,
    get_guidelines_odoo,
    get_client_success_executives_odoo,
    get_service_category_1_options,
    get_service_category_2_options,
    get_all_employees_in_planning,
    find_earliest_available_slot,
    get_companies,
    get_retainer_projects,
    get_retainer_customers,
    get_project_id_by_name,
    update_task_designer,
    get_odoo_connection,
    check_odoo_connection,
    get_available_fields
)
from gmail_integration import get_gmail_service, fetch_recent_emails
from azure_llm import analyze_email
from designer_selector import (
    load_designers,
    suggest_best_designer,
    suggest_best_designer_available,
    filter_designers_by_availability,
    rank_designers_by_skill_match
)

from prezlab_ui import inject_custom_css, header, container, message, progress_steps, scribble, add_logo

# ─── Logging ─────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(name)s [%(levelname)s] %(message)s",
    filename="app.log"
)
logger = logging.getLogger(__name__)


def add_debug_sidebar(debugger: SystemDebugger):
    """
    Add a debug sidebar option to the existing sidebar
    """
    if st.session_state.user['username'] == 'admin':
        st.sidebar.markdown("---")
        st.sidebar.subheader("🐞 Debugging")
        if st.sidebar.button("System Debug Dashboard"):
            # Switch to debug mode
            st.session_state.debug_mode = "system_debug"

def handle_debug_mode(debugger: SystemDebugger):
    """
    Handle the system debug mode rendering
    """
    if st.session_state.get("debug_mode") == "system_debug":
        debugger.streamlit_debug_page()
        # Add a button to return to normal mode
        if st.button("Return to Normal Mode"):
            st.session_state.pop("debug_mode")
            st.rerun()
        return True
    return False

def setup_debugging(main_app):
    """
    Set up debugging for the main Streamlit application
    """
    # Inject debug handlers and get debugger instance
    debugger = inject_debug_page(main_app)
    
    # Modify the sidebar render function to add debug option
    original_render_sidebar = main_app.render_sidebar
    
    def modified_render_sidebar():
        original_render_sidebar()
        add_debug_sidebar(debugger)
    
    main_app.render_sidebar = modified_render_sidebar
    
    return debugger

# Replace environment variables with secrets
ODOO_URL = get_secret("ODOO_URL")
ODOO_DB = get_secret("ODOO_DB")
ODOO_USERNAME = get_secret("ODOO_USERNAME") 
ODOO_PASSWORD = get_secret("ODOO_PASSWORD")


# Set page config
st.set_page_config(
    page_title="Task Management System",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

def validate_session():
    """
    Validates the current session and handles expiry
    
    Returns:
        True if session is valid, False if expired or not logged in
    """
    from session_manager import SessionManager
    
    # CRITICAL FIX: Skip validation if OAuth flow is in progress
    if "code" in st.query_params:
        return True
    
    # Update activity timestamp
    SessionManager.update_activity()
    
    # Check if logged in
    if not st.session_state.get("logged_in", False):
        return False
    
    # Check for session expiry
    if not SessionManager.check_session_expiry():
        return False
    
    # Validate Odoo connection
    if not check_odoo_connection():
        with st.spinner("Reconnecting to Odoo..."):
            uid, models = get_odoo_connection(force_refresh=True)
            if not uid or not models:
                st.error("Lost connection to Odoo. Please log in again.")
                SessionManager.logout()
                return False
    
    return True

# -------------------------------
# SIDEBAR
# -------------------------------
def render_sidebar():
    import xmlrpc.client
    from config import get_secret
    from session_manager import SessionManager

    st.sidebar.title("Task Management")

    # ─── Always‑visible debug info ──────────────────────────────────────────
    st.sidebar.markdown("#### Debug Info:")
    st.sidebar.write("Logged in:", st.session_state.get("logged_in", False))
    st.sidebar.write("Username:", st.session_state.get("user", {}).get("username", "None"))
    # Show session expiry if available
    expiry = st.session_state.get("session_expiry")
    if expiry:
        st.sidebar.write("Session expires at:", expiry.strftime("%Y-%m-%d %H:%M:%S"))
    st.sidebar.markdown("---")

    # ─── Navigation & Auth (only if logged in) ───────────────────────────
    if st.session_state.get("logged_in", False):
        # Navigation
        st.sidebar.subheader("Navigation")
        if st.sidebar.button("Home"):
            SessionManager.clear_flow_data()
            st.rerun()

        # Google Services
        st.sidebar.subheader("Google Services")
        gmail_ok = "google_gmail_creds" in st.session_state
        drive_ok = "google_drive_creds" in st.session_state
        status = "✅ Connected" if (gmail_ok and drive_ok) else "⚠️ Not connected"
        if st.sidebar.button(f"Google Auth ({status})"):
            st.session_state.show_google_auth = True
            st.rerun()

        # Odoo Connection
        st.sidebar.subheader("Connection")
        if st.sidebar.button("Reconnect to Odoo"):
            with st.spinner("Reconnecting..."):
                uid, models = get_odoo_connection(force_refresh=True)
                if uid:
                    st.sidebar.success("Reconnected successfully!")
                else:
                    st.sidebar.error("Failed to reconnect. Check logs.")
        if st.sidebar.button("Logout"):
            SessionManager.logout()
            st.rerun()

        # Admin Tools for 'admin' user
        if st.session_state.user.get("username") == "admin":
            st.sidebar.markdown("---")
            st.sidebar.subheader("Admin Tools")
            if st.sidebar.button("System Debug Dashboard"):
                st.session_state.debug_mode = "system_debug"
                st.rerun()
            if st.sidebar.button("Auth Debug Dashboard"):
                st.session_state.debug_mode = "auth_debug"
                st.rerun()
    else:
        st.sidebar.info("Please log in to access navigation.")

    # ─── Quick Debug (always visible) ─────────────────────────────────────
    st.sidebar.markdown("---")
    st.sidebar.subheader("Quick Debug")
    
    if st.sidebar.button("Test Supabase"):
        # your existing Supabase test code here...
        pass

    if st.sidebar.button("Test Encryption"):
        # your existing Encryption test code here...
        pass

    if st.sidebar.button("Test Odoo Secrets"):
        for key in ["ODOO_URL", "ODOO_DB", "ODOO_USERNAME", "ODOO_PASSWORD"]:
            val = get_secret(key)
            st.sidebar.write(f"**{key}** →", "✅ set" if val else "❌ EMPTY")

    if st.sidebar.button("Test Odoo Connection"):
        url = get_secret("ODOO_URL") + "/xmlrpc/2/common"
        try:
            uid = xmlrpc.client.ServerProxy(url).authenticate(
                get_secret("ODOO_DB"),
                get_secret("ODOO_USERNAME"),
                get_secret("ODOO_PASSWORD"),
                {}
            )
            st.sidebar.success(f"authenticate() → {uid!r}")
        except Exception as e:
            st.sidebar.error(f"RPC error: {type(e).__name__}: {e}")

    st.sidebar.markdown("---")
    st.sidebar.caption("© 2025 Task Management System")


def auth_debug_page():
    """Dashboard for authentication debugging"""
    st.title("Authentication Debug Dashboard")
    
    # Basic Authentication Info
    st.subheader("Authentication Status")
    
    # Display current authentication state
    auth_state = {
        "Logged In": st.session_state.get("logged_in", False),
        "Username": st.session_state.get("user", {}).get("username", "None"),
        "Gmail Auth Complete": st.session_state.get("gmail_auth_complete", False),
        "Drive Auth Complete": st.session_state.get("drive_auth_complete", False),
        "Google Auth Complete": st.session_state.get("google_auth_complete", False),
        "Gmail Credentials Present": "google_gmail_creds" in st.session_state,
        "Drive Credentials Present": "google_drive_creds" in st.session_state
    }
    
    for key, value in auth_state.items():
        st.write(f"**{key}:** {value}")

    # In the auth_debug_page function, add:
    st.subheader("Token Reset")
    if st.button("Reset All OAuth Tokens"):
        from token_storage import reset_user_tokens
        success = reset_user_tokens()
        if success:
            # Also clear local session tokens
            for key in ["google_gmail_creds", "google_drive_creds", 
                        "gmail_auth_complete", "drive_auth_complete", 
                        "google_auth_complete"]:
                if key in st.session_state:
                    del st.session_state[key]
            st.success("All tokens reset successfully. Please re-authenticate.")
        else:
            st.error("Failed to reset tokens. Check logs for details.")
    
    # Supabase Testing
    st.subheader("Supabase Connection Test")
    
    if st.button("Test Supabase Connection"):
        try:
            from token_storage import get_supabase_client
            client = get_supabase_client()
            if client:
                try:
                    response = client.table("oauth_tokens").select("count").limit(1).execute()
                    st.success("✅ Supabase connection successful")
                    st.json(response.data)
                except Exception as e:
                    st.error(f"❌ Table error: {str(e)}")
            else:
                st.error("❌ Supabase client creation failed")
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
    
    # Test encryption
    st.subheader("Encryption Test")
    
    if st.button("Test Encryption System"):
        try:
            from token_storage import get_encryption_key, encrypt_token, decrypt_token
            
            # Get encryption key
            key = get_encryption_key()
            if key:
                st.success(f"✅ Encryption key found (length: {len(key)})")
                
                # Test encryption/decryption
                import time as time_module  # Add this near your other imports
                # Then use:
                test_data = {"test": "data", "time": str(time_module.time())}
                st.write("Test data:", test_data)
                
                encrypted = encrypt_token(test_data)
                if encrypted:
                    st.success(f"✅ Encryption successful")
                    st.text(f"Encrypted data preview: {encrypted[:50]}...")
                    
                    decrypted = decrypt_token(encrypted)
                    if decrypted == test_data:
                        st.success("✅ Decryption successful")
                        st.write("Decrypted data:", decrypted)
                    else:
                        st.error("❌ Decryption failed or mismatched")
                        st.write("Decrypted data:", decrypted)
                else:
                    st.error("❌ Encryption failed")
            else:
                st.error("❌ No encryption key found")
                secrets_keys = list(st.secrets.keys())
                st.write("Available secret keys:", secrets_keys)
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.code(traceback.format_exc())
    
    # Google Auth Actions
    st.subheader("Google Authentication Actions")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Reset Google Auth State"):
            keys = ["google_gmail_creds", "google_drive_creds", 
                    "gmail_auth_complete", "drive_auth_complete", "google_auth_complete"]
            for key in keys:
                if key in st.session_state:
                    del st.session_state[key]
            st.success("Google authentication state reset")
            st.rerun()
    
    with col2:
        if st.button("Start Google Auth Process"):
            st.session_state.show_google_auth = True
            st.rerun()
    
    # Add ability to return to normal mode
    if st.button("Return to Normal Mode"):
        st.session_state.pop("debug_mode", None)
        st.rerun()        
# -------------------------------
# 1) LOGIN PAGE
# -------------------------------

def login_page():
    from session_manager import SessionManager

    # Touch the session so we don't expire mid‑login
    SessionManager.update_activity()
    
    # Add the logo with center positioning for the login page
    add_logo(position="center", width=200)  # Bigger logo for login page

    # Center the form in the middle column
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("Welcome")
        st.subheader("Login to Task Management System")

        with st.form("login_form"):
            username = st.text_input("Username", key="username_input")
            password = st.text_input("Password", type="password", key="password_input")
            submit = st.form_submit_button("Login")

            if not submit:
                return

            # 1) Require both fields
            if not username or not password:
                st.warning("Please enter both username and password.")
                return

            # 2) Validate against multiple user credentials
            authenticated = False
            
            # First check the admin user (backward compatibility)
            valid_admin_user = get_secret("APP_USERNAME", "admin")
            valid_admin_pass = get_secret("APP_PASSWORD", "password")
            if username == valid_admin_user and password == valid_admin_pass:
                authenticated = True
            
            # Check additional users
            for prefix in ["USER_karmel", "USER_abdelrauof"]:
                user = get_secret(f"{prefix}_USERNAME")
                pwd = get_secret(f"{prefix}_PASSWORD")
                if username == user and password == pwd:
                    authenticated = True
                    break
            
            if not authenticated:
                st.error("Invalid credentials. Please try again.")
                return

            # 3) Log into Streamlit session
            SessionManager.login(username, expiry_hours=8)

            # 4) Attempt Odoo authentication exactly once
            try:
                with st.spinner("Connecting to Odoo…"):
                    uid, models = get_odoo_connection(force_refresh=True)

                if not uid:
                    # Explicit error if Odoo.authenticate returned False
                    raise RuntimeError("authenticate() returned False")

                # Success path
                st.success("Login successful!")
                st.rerun()

            except Exception as e:
                # Show full error, log stack trace, and rollback Streamlit login
                st.error(f"Odoo authentication failed: {type(e).__name__}: {e}")
                logger.exception("Odoo login error")
                SessionManager.logout()
                return


# -------------------------------
# 2) REQUEST TYPE SELECTION PAGE
# -------------------------------
def type_selection_page():
    st.title("Task Management Dashboard")
    
    # User greeting
    st.markdown(f"### Welcome, {st.session_state.user['username']}!")
    st.markdown("Select the type of request you want to create.")
    
    col1, col2 = st.columns(2)
    
    with col1:
        with st.container(border=True):
            st.subheader("Via Sales Order")
            st.markdown("For one-time projects or framework tasks with subtasks.")
            if st.button("Create Ad-hoc Request", use_container_width=True):
                st.session_state.form_type = "Via Sales Order"
                st.rerun()
    
    with col2:
        with st.container(border=True):
            st.subheader("Via Project")
            st.markdown("For ongoing projects with recurring tasks.")
            if st.button("Create Retainer Request", use_container_width=True):
                st.session_state.form_type = "Via Project"
                st.rerun()
                
    # Recent activities section
    with st.expander("Recent Activities", expanded=False):
        st.markdown("This section will show recent task activities.")
        # Here you could display recent tasks from Odoo

# -------------------------------
# HELPER: Fetch Sales Order Lines
# -------------------------------
def get_sales_order_lines(models, uid, sales_order_name):
    try:
        so_data = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'sale.order', 'search_read',
            [[['name', '=', sales_order_name]]],
            {'fields': ['order_line']}
        )
        if not so_data:
            return []
        order_line_ids = so_data[0].get('order_line', [])
        if not order_line_ids:
            return []
        lines = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,  # Fixed: Using constants instead of os.getenv
            'sale.order.line', 'read',
            [order_line_ids],
            {'fields': ['id', 'name']}
        )
        return lines
    except Exception as e:
        logger.error(f"Error fetching sales order lines: {e}")
        st.error(f"Error fetching sales order lines. Please try again.")
        return []

# -------------------------------
# 3A) SALES ORDER PAGE (Ad-hoc Step 1)
# -------------------------------
def sales_order_page():
    st.title("Via Sales Order: Select Sales Order")
    
    # Progress bar
    st.progress(50, text="Step 2 of 4: Select Sales Order")
    
    # Navigation
    cols = st.columns([1, 5])
    with cols[0]:
        if st.button("← Back"):
            # Set a flag to indicate we're returning to parent (don't clear the form data)
            st.session_state.pop("company_selection_done", None)
            st.rerun()
    
    # Display selected company
    with st.container(border=True):
        selected_company = st.session_state.get("selected_company", "")
        st.markdown(f"**Selected Company:** {selected_company}")
    
    # Add this code to detect company changes and force refresh
    if "last_company_for_sales_orders" in st.session_state:
        if st.session_state.last_company_for_sales_orders != selected_company:
            # Company changed, force refresh sales orders
            st.session_state.refresh_sales_orders = True
            
    # Update the tracked company
    st.session_state.last_company_for_sales_orders = selected_company

    # Connect to Odoo
    if "odoo_uid" not in st.session_state or "odoo_models" not in st.session_state:
        with st.spinner("Connecting to Odoo..."):
            uid, models = authenticate_odoo()
            if uid and models:
                st.session_state.odoo_uid = uid
                st.session_state.odoo_models = models
            else:
                st.error("Failed to connect to Odoo. Please check your credentials.")
                return
    else:
        uid = st.session_state.odoo_uid
        models = st.session_state.odoo_models

    # Fetch sales orders filtered by company if not already in session
    if "sales_orders" not in st.session_state or st.session_state.get("refresh_sales_orders", False):
        with st.spinner("Fetching sales orders..."):
            orders = get_sales_orders(models, uid, selected_company)
            if orders:
                st.session_state.sales_orders = orders
                st.session_state.refresh_sales_orders = False
            else:
                st.warning(f"No sales orders found for {selected_company} or error fetching orders.")
                st.session_state.sales_orders = []

    # Create form for sales order selection
    with st.form("sales_order_form"):
        st.subheader("Select Sales Order")
        
        sales_order_options = [order['name'] for order in st.session_state.sales_orders]
        selected_sales_order = st.selectbox(
            "Sales Order Number",
            ["(Manual Entry)"] + sales_order_options
        )

        # We're not showing these fields anymore, but we need placeholder 
        # variables that will hold the values later
        parent_sales_order_item = None
        customer = None
        project = None

        submit = st.form_submit_button("Next")
        
        if submit:
            # If a sales order is selected, get the details from Odoo
            if selected_sales_order != "(Manual Entry)":
                details = get_sales_order_details(models, uid, selected_sales_order)
                parent_sales_order_item = details.get('sales_order', selected_sales_order)
                customer = details.get('customer', "")
                project = details.get('project', "")
                
                # For automatic sales order selection, we use the order name itself if no item specified
                if not parent_sales_order_item:
                    parent_sales_order_item = selected_sales_order
            else:
                # For manual entry, we use the selected_sales_order as the item name
                # This removes the validation problem since we're not showing the fields to fill
                parent_sales_order_item = "Manual Entry"
                customer = "Manual Customer"
                project = "Manual Project"
            
            # Save to session state
            st.session_state.parent_sales_order_item = parent_sales_order_item
            st.session_state.customer = customer
            st.session_state.project = project
            
            # Get sales order lines for selected order
            if selected_sales_order != "(Manual Entry)":
                with st.spinner("Fetching sales order lines..."):
                    so_lines = get_sales_order_lines(models, uid, selected_sales_order)
            else:
                so_lines = []
                
            st.session_state.so_items = so_lines
            st.session_state.subtask_index = 0
            st.session_state.adhoc_subtasks = []
            st.session_state.adhoc_sales_order_done = True
            st.success("Sales order information saved. Proceeding to parent task details.")
            st.rerun()

# -------------------------------
# 3B) AD-HOC PARENT TASK PAGE (Ad-hoc Step 2)
# -------------------------------
def adhoc_parent_task_page():
    st.title("Via Sales Order: Parent Task")
    
    # Progress bar
    st.progress(80, text="Step 4 of 5: Parent Task Details")
    
    # Clear the return flag if it exists
    st.session_state.pop("returning_to_parent", False) if "returning_to_parent" in st.session_state else None

    # Navigation
    cols = st.columns([1, 5])
    with cols[0]:
        if st.button("← Back"):
            st.session_state.pop("adhoc_sales_order_done", None)
            st.rerun()

    # Display current selection
    with st.container(border=True):
        selected_company = st.session_state.get("selected_company", "")
        parent_sales_order_item = st.session_state.get("parent_sales_order_item", "")
        customer = st.session_state.get("customer", "")
        project = st.session_state.get("project", "")
        
        st.markdown("### Current Selection")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f"**Company:** {selected_company}")
        with col2:
            st.markdown(f"**Sales Order:** {parent_sales_order_item}")
        with col3:
            st.markdown(f"**Customer:** {customer}")
        with col4:
            st.markdown(f"**Project:** {project}")
    
    uid = st.session_state.odoo_uid
    models = st.session_state.odoo_models

    # Get email analysis results if available
    email_analysis = st.session_state.get("email_analysis", {})
    email_analysis_skipped = st.session_state.get("email_analysis_skipped", True)

    # Form for parent task details
    with st.form("parent_task_form"):
        st.subheader("Parent Task Details")
        
        # Basic Information
        with st.container():
            # Use email analysis results as default values if available
            default_title = ""
            if email_analysis and isinstance(email_analysis, dict) and not email_analysis_skipped:
                services = email_analysis.get("services", "")
                if services:
                    default_title = f"Task for {services}"
            
            parent_task_title = st.text_input("Parent Task Title", 
                                             value=default_title,
                                             help="Enter a descriptive title for the parent task")
            
            col1, col2 = st.columns(2)
            with col1:
                target_language_options = get_target_languages_odoo(models, uid)
                # Default target language from email analysis
                default_target_lang_idx = 0
                if email_analysis and isinstance(email_analysis, dict) and not email_analysis_skipped:
                    target_lang = email_analysis.get("target_language", "")
                    if target_lang and target_language_options:
                        for i, lang in enumerate(target_language_options):
                            if target_lang.lower() in lang.lower():
                                default_target_lang_idx = i
                                break
                
                target_language_parent = st.selectbox(
                    "Target Language", 
                    target_language_options if target_language_options else [""],
                    index=default_target_lang_idx,
                    help="Select the target language for this task"
                )

            with col2:
                client_success_exec_options = get_client_success_executives_odoo(models, uid)
                if client_success_exec_options:
                    exec_options = [(user['id'], user['name']) for user in client_success_exec_options]
                    client_success_executive = st.selectbox(
                        "Client Success Executive", 
                        options=exec_options, 
                        format_func=lambda x: x[1],
                        help="Select the responsible client success executive"
                    )
                else:
                    client_success_executive = st.text_input("Client Success Executive")
        
        # Guidelines
        with st.expander("Guidelines", expanded=False):
            guidelines_options = get_guidelines_odoo(models, uid)
            if guidelines_options:
                # Use format_func to display the name while storing the tuple
                guidelines_parent = st.selectbox(
                    "Guidelines", 
                    options=guidelines_options,
                    format_func=lambda x: x[1]  # Display the name part
                )
            else:
                st.error("No guidelines found. This field is required.")
                guidelines_parent = None
        # Dates
        st.subheader("Task Timeline")
        col1, col2, col3 = st.columns(3)
        with col1:
            request_receipt_date = st.date_input("Request Receipt Date", value=date.today())
            request_receipt_time = st.time_input("Request Receipt Time", value=datetime.now().time())
        with col2:
            client_due_date_parent = st.date_input("Client Due Date", value=date.today() + pd.Timedelta(days=7))
        with col3:
            internal_due_date = st.date_input("Internal Due Date", value=date.today() + pd.Timedelta(days=5))
        
        # Combine date and time
        request_receipt_dt = datetime.combine(request_receipt_date, request_receipt_time)
        
        # Description
        st.subheader("Description")
        
        # Use requirements from email analysis as starting point for description
        default_description = ""
        if email_analysis and isinstance(email_analysis, dict) and not email_analysis_skipped:
            requirements = email_analysis.get("requirements", "")
            if requirements:
                default_description = f"Requirements from client email:\n{requirements}"
                
                # Add other relevant information
                services = email_analysis.get("services", "")
                if services:
                    default_description += f"\n\nRequested Services: {services}"
                    
                deadline = email_analysis.get("deadline", "")
                if deadline:
                    default_description += f"\n\nClient Requested Deadline: {deadline}"
        
        parent_description = st.text_area("Task Description", 
                                         value=default_description,
                                         height=150, 
                                         help="Enter any additional details for this task")
        # Submit button
        submit = st.form_submit_button("Next: Add Subtasks")
        
        if submit:
            # Validate inputs
            if not parent_task_title:
                st.error("Please enter a parent task title.")
                return
                
            # Save to session state
            st.session_state.adhoc_parent_task_title = parent_task_title
            st.session_state.adhoc_target_language = target_language_parent
            st.session_state.adhoc_guidelines = guidelines_parent
            st.session_state.adhoc_client_success_exec = client_success_executive
            st.session_state.adhoc_request_receipt_dt = request_receipt_dt
            st.session_state.adhoc_client_due_date_parent = client_due_date_parent
            st.session_state.adhoc_internal_due_date = internal_due_date
            st.session_state.adhoc_parent_description = parent_description
            st.session_state.adhoc_parent_input_done = True
            
            st.success("Parent task details saved. Proceeding to subtasks.")
            st.rerun()

# -------------------------------
# 3C) AD-HOC SUBTASK PAGE (Ad-hoc Step 3)
# -------------------------------
def adhoc_subtask_page():
    st.title("Via Sales Order: Subtasks")
    
    # Progress bar
    st.progress(100, text="Step 5 of 5: Subtask Details")
    
    # Navigation
    cols = st.columns([1, 5])
    with cols[0]:
        if st.button("← Back"):
            st.session_state.pop("adhoc_parent_input_done", None)
            st.rerun()

    uid = st.session_state.odoo_uid
    models = st.session_state.odoo_models

    # Get parent task information
    parent_data = {
        "selected_company": st.session_state.get("selected_company", ""),
        "parent_sales_order_item": st.session_state.get("parent_sales_order_item", ""),
        "parent_task_title": st.session_state.get("adhoc_parent_task_title", ""),
        "customer": st.session_state.get("customer", ""),
        "project": st.session_state.get("project", ""),
        "target_language_parent": st.session_state.get("adhoc_target_language", ""),
        "guidelines_parent": st.session_state.get("adhoc_guidelines", ""),
        "client_success_executive": st.session_state.get("adhoc_client_success_exec", ""),
        "request_receipt_dt": st.session_state.get("adhoc_request_receipt_dt", datetime.now()),
        "client_due_date_parent": st.session_state.get("adhoc_client_due_date_parent", date.today()),
        "internal_due_date": st.session_state.get("adhoc_internal_due_date", date.today()),
        "parent_description": st.session_state.get("adhoc_parent_description", "")
    }
    
    # Display parent task summary
    with st.expander("Parent Task Summary", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**Company:** {parent_data['selected_company']}")
            st.markdown(f"**Parent Task:** {parent_data['parent_task_title']}")
            st.markdown(f"**Sales Order:** {parent_data['parent_sales_order_item']}")
            st.markdown(f"**Customer:** {parent_data['customer']}")
            st.markdown(f"**Project:** {parent_data['project']}")
            st.markdown(f"**Target Language:** {parent_data['target_language_parent']}")
        with col2:
            st.markdown(f"**Client Success Executive:** {parent_data['client_success_executive'][1] if isinstance(parent_data['client_success_executive'], tuple) else parent_data['client_success_executive']}")
            st.markdown(f"**Request Receipt:** {parent_data['request_receipt_dt'].strftime('%Y-%m-%d %H:%M')}")
            st.markdown(f"**Client Due Date:** {parent_data['client_due_date_parent']}")
            st.markdown(f"**Internal Due Date:** {parent_data['internal_due_date']}")
        
        st.markdown(f"**Description:** {parent_data['parent_description']}")

    # Rest of the function remains the same...
    
    # Get current subtask index and sales order items list
    idx = st.session_state.get("subtask_index", 0)
    so_items = st.session_state.get("so_items", [])
    
    # Display subtasks in progress
    if st.session_state.adhoc_subtasks:
        with st.container(border=True):
            st.subheader("Subtasks Added")
            for i, task in enumerate(st.session_state.adhoc_subtasks):
                st.markdown(f"**Subtask {i+1}:** {task['subtask_title']} (Due: {task['client_due_date_subtask']})")
    
    # Check if we have more subtasks to add
    if idx >= len(so_items):
        st.warning("No more sales order items available for subtasks.")
        
        # Allow adding a custom subtask if needed
        with st.expander("Add Custom Subtask", expanded=True):
            with st.form("custom_subtask_form"):
                custom_subtask_title = st.text_input("Custom Subtask Title")
                
                col1, col2 = st.columns(2)
                with col1:
                    service_category_1_options = get_service_category_1_options(models, uid)
                    service_category_1 = st.selectbox(
                        "Service Category 1", 
                        [opt[1] if isinstance(opt, list) and len(opt) > 1 else opt for opt in service_category_1_options] if service_category_1_options else [""]
                    )
                    no_of_design_units_sc1 = st.number_input("Total No. of Design Units (SC1)", min_value=0, step=1)
                
                with col2:
                    service_category_2_options = get_service_category_2_options(models, uid)
                    service_category_2 = st.selectbox(
                        "Service Category 2", 
                        [opt[1] if isinstance(opt, list) and len(opt) > 1 else opt for opt in service_category_2_options] if service_category_2_options else [""]
                    )
                    no_of_design_units_sc2 = st.number_input("Total No. of Design Units (SC2)", min_value=0, step=1)
                
                client_due_date_subtask = st.date_input("Client Due Date (Subtask)", value=date.today() + pd.Timedelta(days=5))
                
                add_custom = st.form_submit_button("Add Custom Subtask")
                
                if add_custom:
                    if not custom_subtask_title:
                        st.error("Please enter a subtask title.")
                        return
                        
                    new_subtask = {
                        "line_id": None,
                        "line_name": "Custom",
                        "subtask_title": custom_subtask_title,
                        "service_category_1": service_category_1,
                        "no_of_design_units_sc1": no_of_design_units_sc1,
                        "service_category_2": service_category_2,
                        "no_of_design_units_sc2": no_of_design_units_sc2,
                        "client_due_date_subtask": str(client_due_date_subtask)
                    }
                    st.session_state.adhoc_subtasks.append(new_subtask)
                    st.success(f"Added custom subtask: {custom_subtask_title}")
                    st.rerun()
        
        # Show submission button
        if st.button("Submit All Tasks", use_container_width=True, type="primary"):
            with st.spinner("Creating tasks in Odoo..."):
                finalize_adhoc_subtasks()
        
        return
    
    # Get email analysis results if available
    email_analysis = st.session_state.get("email_analysis", {})
    email_analysis_skipped = st.session_state.get("email_analysis_skipped", True)

    # Current sales order line for this subtask
    current_line = so_items[idx]
    line_name = current_line.get("name", f"Line #{idx+1}")
    
    # Subtask form
    # Current sales order line for this subtask
    current_line = so_items[idx] if idx < len(so_items) else {}
    line_name = current_line.get("name", f"Line #{idx+1}") if current_line else f"Subtask #{idx+1}"
    
    # Subtask form
    with st.form(f"subtask_form_{idx}"):
        st.subheader(f"Subtask for Sales Order Line: {line_name}")
        
        # Default title from services in email analysis
        default_title = f"Subtask for {line_name}"
        if email_analysis and isinstance(email_analysis, dict) and not email_analysis_skipped and idx == 0:  # Only use for first subtask
            services = email_analysis.get("services", "")
            if services:
                default_title = f"Subtask: {services}"
                
        subtask_title = st.text_input("Subtask Title", value=default_title)
        
        col1, col2 = st.columns(2)
        with col1:
            service_category_1_options = get_service_category_1_options(models, uid)
            if service_category_1_options:
                service_category_1 = st.selectbox(
                    "Service Category 1", 
                    options=service_category_1_options,
                    format_func=lambda x: x[1] if isinstance(x, tuple) and len(x) > 1 else str(x)
                )
            else:
                # Fallback to text input with warning
                st.warning("No service categories found. Manual entry not recommended.")
                service_category_1_text = st.text_input("Service Category 1 (manual)")
                # Create a dummy tuple with -1 as ID to indicate this is not a valid ID
                service_category_1 = (-1, service_category_1_text) if service_category_1_text else None
            
            no_of_design_units_sc1 = st.number_input("Total No. of Design Units (SC1)", min_value=0, step=1)

        with col2:
            service_category_2_options = get_service_category_2_options(models, uid)
            if service_category_2_options:
                service_category_2 = st.selectbox(
                    "Service Category 2", 
                    options=service_category_2_options,
                    format_func=lambda x: x[1] if isinstance(x, tuple) and len(x) > 1 else str(x)
                )
            else:
                # Fallback to text input with warning
                st.warning("No service categories found. Manual entry not recommended.")
                service_category_2_text = st.text_input("Service Category 2 (manual)")
                # Create a dummy tuple with -1 as ID to indicate this is not a valid ID
                service_category_2 = (-1, service_category_2_text) if service_category_2_text else None
            
            no_of_design_units_sc2 = st.number_input("Total No. of Design Units (SC2)", min_value=0, step=1)

        
        client_due_date_subtask = st.date_input("Client Due Date (Subtask)", value=date.today() + pd.Timedelta(days=5))
        
        # Submit options
        col1, col2 = st.columns(2)
        with col1:
            next_subtask = st.form_submit_button("Save & Next Subtask")
        with col2:
            if idx == len(so_items) - 1:
                finish_all = st.form_submit_button("Finish & Submit All")
            else:
                finish_all = False
        
        if next_subtask or finish_all:
            # Validate input
            if not subtask_title:
                st.error("Please enter a subtask title.")
                return
                
            # Create new subtask
            new_subtask = {
                "line_id": current_line.get("id"),
                "line_name": line_name,
                "subtask_title": subtask_title,
                "service_category_1": service_category_1,
                "no_of_design_units_sc1": no_of_design_units_sc1,
                "service_category_2": service_category_2,
                "no_of_design_units_sc2": no_of_design_units_sc2,
                "client_due_date_subtask": str(client_due_date_subtask)
            }
            
            # Add to session state
            st.session_state.adhoc_subtasks.append(new_subtask)
            st.session_state.subtask_index = idx + 1
            
            # If finishing, submit all; otherwise, continue to next subtask
            if finish_all:
                with st.spinner("Creating tasks in Odoo..."):
                    finalize_adhoc_subtasks()
            else:
                st.success(f"Subtask saved: {subtask_title}")
                st.rerun()

# -------------------------------
# Finalize: Create Parent Task & Subtasks in Odoo
# -------------------------------
def finalize_adhoc_subtasks():
    from session_manager import SessionManager
    SessionManager.update_activity()
    
    # Import modules only when needed
    from google_drive import create_folder_structure
    uid = st.session_state.odoo_uid
    models = st.session_state.odoo_models

    # Get parent task data
    selected_company = st.session_state.get("selected_company", "")
    parent_sales_order_item = st.session_state.get("parent_sales_order_item", "")
    parent_task_title = st.session_state.get("adhoc_parent_task_title", "")
    customer = st.session_state.get("customer", "")
    project_name = st.session_state.get("project", "")
    target_language_parent = st.session_state.get("adhoc_target_language", "")
    guidelines_parent = st.session_state.get("adhoc_guidelines", "")
    client_success_executive = st.session_state.get("adhoc_client_success_exec", "")
    request_receipt_dt = st.session_state.get("adhoc_request_receipt_dt", datetime.now())
    client_due_date_parent = st.session_state.get("adhoc_client_due_date_parent", date.today())
    internal_due_date = st.session_state.get("adhoc_internal_due_date", date.today())
    parent_description = st.session_state.get("adhoc_parent_description", "")

    try:
        # Get project ID
        project_id = get_project_id_by_name(models, uid, project_name)
        if not project_id:
            st.error(f"Could not find project with name: {project_name}")
            return
            
        # Ensure project_id is integer
        if not isinstance(project_id, int):
            try:
                project_id = int(project_id)
            except (ValueError, TypeError) as e:
                st.error(f"Invalid project ID format: {e}")
                logger.error(f"Invalid project ID format: {project_id}, error: {e}")
                return
        
        # Handle user ID
        user_id = client_success_executive[0] if isinstance(client_success_executive, tuple) else client_success_executive
        if not isinstance(user_id, int):
            try:
                user_id = int(user_id)
            except (ValueError, TypeError) as e:
                st.error(f"Invalid user ID format: {e}")
                logger.error(f"Invalid user ID format: {user_id}, error: {e}")
                return
        
        # Handle guidelines_id - make sure it's an integer
        guidelines_id = guidelines_parent[0] if isinstance(guidelines_parent, tuple) else None
        
        # Create parent task with only fields that exist
        parent_task_data = {
            "name": f"Ad-hoc Parent: {parent_task_title}",
            "project_id": project_id,
            "user_ids": [(6, 0, [user_id])],  # This is the correct format for many2many fields
            "description": f"Company: {selected_company}\nSales Order Item: {parent_sales_order_item}\nCustomer: {customer}\nProject: {project_name}\n{parent_description}"
        }
        
        # Add optional fields if they exist and have values
        if target_language_parent:
            parent_task_data["x_studio_target_language"] = target_language_parent
            
        if guidelines_id:
            parent_task_data["x_studio_guidelines"] = guidelines_id
        
        # Format dates correctly to avoid the microseconds issue
        if request_receipt_dt:
            parent_task_data["x_studio_request_receipt_date_time"] = request_receipt_dt.strftime("%Y-%m-%d %H:%M:%S")
            
        if client_due_date_parent:
            parent_task_data["x_studio_client_due_date_3"] = client_due_date_parent.strftime("%Y-%m-%d")
            
        if internal_due_date:
            parent_task_data["x_studio_internal_due_date_1"] = internal_due_date.strftime("%Y-%m-%d")
        
        # Create parent task in Odoo
        parent_task_id = create_odoo_task(parent_task_data)
        if not parent_task_id:
            st.error("Failed to create parent task in Odoo.")
            return
            
        st.success(f"Created Parent Task in Odoo (ID: {parent_task_id})")
        
        # Create Google Drive folder structure for this parent task
        with st.spinner("Creating Google Drive folder structure for task..."):
            # Sanitize folder name (replace characters not allowed in file names)
            folder_name = f"{parent_task_title} - {parent_task_id}"
            folder_name = folder_name.replace('/', '-').replace('\\', '-')
            
            # Create folder structure with subfolders
            from google_drive import create_folder_structure
            folder_structure = create_folder_structure(
                folder_name, 
                subfolders=["MATERIAL", "DELIVERABLE"]
            )
            
            if folder_structure:
                # Store main folder info in session state
                st.session_state.drive_folder_id = folder_structure['main_folder_id']
                st.session_state.drive_folder_link = folder_structure['main_folder_link']
                st.session_state.folder_structure = folder_structure
                
                # Update the parent task with the folder links
                try:
                    # Create a nicely formatted description with folder links
                    updated_description = f"{parent_description}\n\n"
                    updated_description += f"📁 **Google Drive Folders:**\n"
                    updated_description += f"- Main Folder: {folder_structure['main_folder_url']}\n"
                    
                    # Add subfolder links if available
                    for subfolder_name, subfolder_info in folder_structure['subfolders'].items():
                        updated_description += f"- {subfolder_name}: {subfolder_info['url']}\n"
                    
                    models.execute_kw(
                        ODOO_DB, uid, ODOO_PASSWORD,
                        'project.task', 'write',
                        [[parent_task_id], {'description': updated_description}]
                    )
                    logger.info(f"Updated task {parent_task_id} with Drive folder structure links")
                except Exception as e:
                    logger.warning(f"Could not update task with folder links: {e}")
                    
                st.success(f"Created folder structure with MATERIAL and DELIVERABLE subfolders")
            else:
                st.warning("Could not create Google Drive folder. Please check logs for details.")
        # Create subtasks
        subtasks = st.session_state.adhoc_subtasks
        created_subtasks = []
        
        for i, sub in enumerate(subtasks):
            # Create base subtask data
            subtask_data = {
                "name": f"Ad-hoc Subtask: {sub['subtask_title']}",
                "project_id": project_id,
                "user_ids": [(6, 0, [user_id])],
                "description": f"Company: {selected_company}\nSubtask for Sales Order Line: {sub['line_name']}",
                # Set the parent-child relationship - use parent_id, not x_studio_sub_task_1
                "parent_id": parent_task_id
            }
            
            # Add optional fields if they exist and have values
            if target_language_parent:
                subtask_data["x_studio_target_language"] = target_language_parent
                
            if guidelines_id:
                subtask_data["x_studio_guidelines"] = guidelines_id
            
            # Format dates correctly
            if request_receipt_dt:
                subtask_data["x_studio_request_receipt_date_time"] = request_receipt_dt.strftime("%Y-%m-%d %H:%M:%S")
                
            # Use the subtask-specific due date
            if "client_due_date_subtask" in sub and sub["client_due_date_subtask"]:
                # Parse the date from string if needed
                if isinstance(sub["client_due_date_subtask"], str):
                    due_date = datetime.strptime(sub["client_due_date_subtask"], "%Y-%m-%d").date()
                else:
                    due_date = sub["client_due_date_subtask"]
                subtask_data["x_studio_client_due_date_3"] = due_date.strftime("%Y-%m-%d")
                
            if internal_due_date:
                subtask_data["x_studio_internal_due_date_1"] = internal_due_date.strftime("%Y-%m-%d")
            
            # Add service categories if they exist
            if "service_category_1" in sub and sub["service_category_1"]:
                # If service_category_1 is a tuple with ID and name, use the ID
                if isinstance(sub["service_category_1"], tuple) and len(sub["service_category_1"]) > 1:
                    # Only use the ID if it's valid (not -1)
                    if sub["service_category_1"][0] != -1:
                        subtask_data["x_studio_service_category_1"] = sub["service_category_1"][0]
                    else:
                        logger.warning(f"Skipping invalid service_category_1 ID: {sub['service_category_1']}")
                # If it's already an ID, use it directly
                elif isinstance(sub["service_category_1"], int):
                    subtask_data["x_studio_service_category_1"] = sub["service_category_1"]
                else:
                    logger.warning(f"Skipping service_category_1 as it's not in the expected format: {sub['service_category_1']}")

            # Similar handling for service_category_2
            if "service_category_2" in sub and sub["service_category_2"]:
                if isinstance(sub["service_category_2"], tuple) and len(sub["service_category_2"]) > 1:
                    # Only use the ID if it's valid (not -1)
                    if sub["service_category_2"][0] != -1:
                        subtask_data["x_studio_service_category_2"] = sub["service_category_2"][0]
                    else:
                        logger.warning(f"Skipping invalid service_category_2 ID: {sub['service_category_2']}")
                elif isinstance(sub["service_category_2"], int):
                    subtask_data["x_studio_service_category_2"] = sub["service_category_2"]
                else:
                    logger.warning(f"Skipping service_category_2 as it's not in the expected format: {sub['service_category_2']}")
            
            # Add design units if applicable
            if "no_of_design_units_sc1" in sub and sub["no_of_design_units_sc1"]:
                subtask_data["x_studio_total_no_of_design_units_sc1"] = sub["no_of_design_units_sc1"]
            
            if "no_of_design_units_sc2" in sub and sub["no_of_design_units_sc2"]:
                subtask_data["x_studio_total_no_of_design_units_sc2"] = sub["no_of_design_units_sc2"]
            
            # Create subtask in Odoo
            subtask_id = create_odoo_task(subtask_data)
            if subtask_id:
                created_subtasks.append(subtask_id)
                st.success(f"Created Subtask {i+1} in Odoo (ID: {subtask_id})")
            else:
                st.error(f"Failed to create subtask {i+1} in Odoo.")
        
        # After tasks are created successfully
        if created_subtasks:
            st.success(f"Successfully created {len(created_subtasks)} subtasks!")
            
            # Store created tasks and parent task ID in session state for designer selection
            created_tasks = []
            for subtask_id in created_subtasks:
                # Fetch the task details from Odoo
                task_details = models.execute_kw(
                    ODOO_DB, uid, ODOO_PASSWORD,
                    'project.task', 'read',
                    [[subtask_id]],
                    {'fields': ['id', 'name', 'description', 'x_studio_service_category_1', 
                            'x_studio_service_category_2', 'x_studio_target_language',
                            'x_studio_client_due_date_3', 'date_deadline']}
                )[0]
                created_tasks.append(task_details)
            
            # Store in session state
            st.session_state.created_tasks = created_tasks
            st.session_state.parent_task_id = parent_task_id
            st.session_state.designer_selection = True
            
            # Summary container
            with st.container(border=True):
                st.subheader("Task Creation Summary")
                st.markdown(f"**Parent Task:** {parent_task_title} (ID: {parent_task_id})")
                st.markdown(f"**Subtasks Created:** {len(created_subtasks)}")
                st.markdown(f"**Company:** {selected_company}")
                st.markdown(f"**Project:** {project_name}")
                st.markdown(f"**Client:** {customer}")
                
                
                if 'drive_folder_id' in st.session_state and 'drive_folder_link' in st.session_state:
                    st.markdown(f"**📁 Main Folder:** [Open Folder]({st.session_state.drive_folder_link})")
                    
                    # If we have subfolder information, display those links too
                    if 'folder_structure' in st.session_state and 'subfolders' in st.session_state.folder_structure:
                        for subfolder_name, subfolder_info in st.session_state.folder_structure['subfolders'].items():
                            st.markdown(f"**📁 {subfolder_name}:** [Open Folder]({subfolder_info['link']})")


                # Display a message to help user understand what's happening
                st.info("Click the button below to proceed to designer selection, or you can view the task details in Odoo.")
                st.session_state.designer_selection = True  # This flag is already being checked elsewhere
                st.rerun()



    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
        logger.error(f"Error in finalize_adhoc_subtasks: {e}", exc_info=True)

# -------------------------------
# RETAINER FLOW
# -------------------------------
def retainer_parent_task_page():
    st.title("Via Project: Parent Task")
    
    # Progress bar
    st.progress(80, text="Step 3 of 4: Parent Task Details")
    
    # Navigation
    cols = st.columns([1, 5])
    with cols[0]:
        if st.button("← Back"):
            st.session_state.returning_to_company = True
            st.rerun()
            
    # Display selected company
    with st.container(border=True):
        selected_company = st.session_state.get("selected_company", "")
        st.markdown(f"**Selected Company:** {selected_company}")
    
    # Connect to Odoo if not already connected
    if "odoo_uid" not in st.session_state or "odoo_models" not in st.session_state:
        with st.spinner("Connecting to Odoo..."):
            uid, models = authenticate_odoo()
            if uid and models:
                st.session_state.odoo_uid = uid
                st.session_state.odoo_models = models
            else:
                st.error("Failed to connect to Odoo. Please check your credentials.")
                return
    else:
        uid = st.session_state.odoo_uid
        models = st.session_state.odoo_models

    # Form for retainer parent task
    with st.form("retainer_parent_form"):
        st.subheader("Retainer Project Fields")
        
        # Project and customer selection
        col1, col2 = st.columns(2)
        with col1:
            retainer_project_options = get_retainer_projects(models, uid, selected_company)
            if retainer_project_options:
                parent_project = st.selectbox("Project", retainer_project_options)
            else:
                parent_project = st.text_input("Project")
        
        with col2:
            retainer_customer_options = get_retainer_customers(models, uid)
            if retainer_customer_options:
                retainer_customer = st.selectbox("Customer", retainer_customer_options)
            else:
                retainer_customer = st.text_input("Customer")
        
        parent_task_title = st.text_input("Parent Task Title")
        
        # Language and executive
        col1, col2 = st.columns(2)
        with col1:
            target_language_options = get_target_languages_odoo(models, uid)
            if target_language_options:
                retainer_target_language = st.selectbox("Target Language", target_language_options)
            else:
                retainer_target_language = st.text_input("Target Language")
        
        with col2:
            client_success_exec_options = get_client_success_executives_odoo(models, uid)
            if client_success_exec_options:
                exec_options = [(user['id'], user['name']) for user in client_success_exec_options]
                retainer_client_success_exec = st.selectbox("Client Success Executive", exec_options, format_func=lambda x: x[1])
            else:
                retainer_client_success_exec = st.text_input("Client Success Executive")
        
        # Guidelines
        with st.expander("Guidelines", expanded=False):
            guidelines_options = get_guidelines_odoo(models, uid)
            if guidelines_options:
                # Use format_func to display only the name while storing the tuple
                retainer_guidelines = st.selectbox(
                    "Guidelines", 
                    options=guidelines_options,
                    format_func=lambda x: x[1]  # Display the name part
                )
            else:
                retainer_guidelines = st.text_area("Guidelines", height=100)
        
        # Dates
        st.subheader("Dates (Retainer Parent Task)")
        
        col1, col2 = st.columns(2)
        with col1:
            retainer_request_receipt_date = st.date_input("Request Receipt Date", value=date.today())
            retainer_request_receipt_time = st.time_input("Request Receipt Time", value=datetime.now().time())
        
        with col2:
            retainer_internal_due_date = st.date_input("Internal Due Date", value=date.today() + pd.Timedelta(days=5))
            retainer_internal_due_time = st.time_input("Internal Due Time", value=time(17, 0))  # 5:00 PM default
        
        # Combine date and time
        retainer_request_receipt_dt = datetime.combine(retainer_request_receipt_date, retainer_request_receipt_time)
        retainer_internal_dt = datetime.combine(retainer_internal_due_date, retainer_internal_due_time)
        
        # Submit button
        submit = st.form_submit_button("Next: Add Subtask")
        
        if submit:
            # Validate input
            if not parent_project or not parent_task_title or not retainer_customer:
                st.error("Please fill in all required fields.")
                return
                
            # Save to session state
            st.session_state.retainer_project = parent_project
            st.session_state.retainer_parent_task_title = parent_task_title
            st.session_state.retainer_customer = retainer_customer
            st.session_state.retainer_target_language = retainer_target_language
            st.session_state.retainer_guidelines = retainer_guidelines
            st.session_state.retainer_client_success_exec = retainer_client_success_exec
            st.session_state.retainer_request_receipt_dt = retainer_request_receipt_dt
            st.session_state.retainer_internal_dt = retainer_internal_dt
            st.session_state.retainer_parent_input_done = True
            
            st.success("Parent task details saved. Proceeding to subtask.")
            st.rerun()

def retainer_subtask_page():
    st.title("Via Project: Subtask")
    
    # Progress bar
    st.progress(100, text="Step 4 of 4: Subtask Details")
    
    # Navigation
    cols = st.columns([1, 5])
    with cols[0]:
        if st.button("← Back"):
            st.session_state.pop("retainer_parent_input_done", None)
            st.rerun()

    uid = st.session_state.odoo_uid
    models = st.session_state.odoo_models

    # Get parent task information
    selected_company = st.session_state.get("selected_company", "")
    parent_project_name = st.session_state.get("retainer_project", "")
    parent_task_title = st.session_state.get("retainer_parent_task_title", "")
    retainer_customer = st.session_state.get("retainer_customer", "")
    retainer_target_language = st.session_state.get("retainer_target_language", "")
    retainer_guidelines = st.session_state.get("retainer_guidelines", "")
    retainer_client_success_exec = st.session_state.get("retainer_client_success_exec", "")
    retainer_request_receipt_dt = st.session_state.get("retainer_request_receipt_dt", datetime.now())
    retainer_internal_dt = st.session_state.get("retainer_internal_dt", datetime.now())

    # Display parent task summary
    with st.container(border=True):
        st.subheader("Parent Task Summary")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**Company:** {selected_company}")
            st.markdown(f"**Project:** {parent_project_name}")
            st.markdown(f"**Parent Task:** {parent_task_title}")
            st.markdown(f"**Customer:** {retainer_customer}")
        with col2:
            st.markdown(f"**Target Language:** {retainer_target_language}")
            st.markdown(f"**Request Receipt:** {retainer_request_receipt_dt.strftime('%Y-%m-%d %H:%M')}")
            st.markdown(f"**Internal Due Date:** {retainer_internal_dt.strftime('%Y-%m-%d %H:%M')}")

    # Subtask form
    with st.form("retainer_subtask_form"):
        st.subheader("Subtask Details")
        
        subtask_title = st.text_input("Subtask Title")
        
        # Service categories
        col1, col2 = st.columns(2)
        with col1:
            service_category_1_options = get_service_category_1_options(models, uid)
            if service_category_1_options:
                retainer_service_category_1 = st.selectbox(
                    "Service Category 1", 
                    options=service_category_1_options,
                    format_func=lambda x: x[1] if isinstance(x, tuple) and len(x) > 1 else str(x)
                )
            else:
                retainer_service_category_1 = st.text_input("Service Category 1")
            
            no_of_design_units_sc1 = st.number_input("No. of Design Units SC1", min_value=0, step=1)
        
        with col2:
            service_category_2_options = get_service_category_2_options(models, uid)
            if service_category_2_options:
                retainer_service_category_2 = st.selectbox(
                    "Service Category 2", 
                    options=service_category_2_options,
                    format_func=lambda x: x[1] if isinstance(x, tuple) and len(x) > 1 else str(x)
                )
            else:
                retainer_service_category_2 = st.text_input("Service Category 2")
            
            no_of_design_units_sc2 = st.number_input("No. of Design Units SC2", min_value=0, step=1)
        
        # Due date
        retainer_client_due_date_subtask = st.date_input("Client Due Date (Subtask)")
        
        # Designer suggestion and submit buttons
        col1, col2 = st.columns(2)

        with col2:
            submit_request = st.form_submit_button("Submit Request")
        
        
        # SUBMIT BUTTON HANDLER - Replace the existing submit_request button handler
        if submit_request:
            # Validate inputs
            if not subtask_title:
                st.error("Please enter a subtask title.")
                return
                
            # Get project ID
            project_id = get_project_id_by_name(models, uid, parent_project_name)
            if not project_id:
                st.error(f"Could not find a project with name: {parent_project_name}")
                return
                
            # Ensure project_id is integer
            if not isinstance(project_id, int):
                try:
                    project_id = int(project_id)
                except (ValueError, TypeError) as e:
                    st.error(f"Invalid project ID format: {e}")
                    logger.error(f"Invalid project ID format: {project_id}, error: {e}")
                    return
            
            # Handle user ID
            user_id = retainer_client_success_exec[0] if isinstance(retainer_client_success_exec, tuple) else retainer_client_success_exec
            if not isinstance(user_id, int):
                try:
                    user_id = int(user_id)
                except (ValueError, TypeError) as e:
                    st.error(f"Invalid user ID format: {e}")
                    logger.error(f"Invalid user ID format: {user_id}, error: {e}")
                    return
            
            # Find partner_id for customer if available
            partner_id = None
            try:
                partners = models.execute_kw(
                    ODOO_DB, uid, ODOO_PASSWORD,
                    'res.partner', 'search_read',
                    [[['name', '=', retainer_customer]]],
                    {'fields': ['id']}
                )
                if partners:
                    partner_id = partners[0]['id']
                    logger.info(f"Found partner_id {partner_id} for customer {retainer_customer}")
            except Exception as e:
                logger.warning(f"Could not find partner_id for customer {retainer_customer}: {e}")
            
            try:
                # STEP 1: CREATE PARENT TASK
                with st.spinner("Creating parent task in Odoo..."):
                    # Create parent task with only fields that exist
                    parent_task_data = {
                        "name": f"Retainer Parent: {parent_project_name} - {retainer_customer}",
                        "project_id": project_id,
                        "user_ids": [(6, 0, [user_id])],  # This is the correct format for many2many fields
                        "description": f"Company: {selected_company}\nCustomer: {retainer_customer}\nProject: {parent_project_name}"
                    }
                    
                    # Add partner_id if found
                    if partner_id:
                        parent_task_data["partner_id"] = partner_id
                    
                    # Add optional fields if they exist and have values
                    if retainer_target_language:
                        parent_task_data["x_studio_target_language"] = retainer_target_language
                    
                    # Handle guidelines properly - extract ID from tuple if applicable
                    if retainer_guidelines:
                        if isinstance(retainer_guidelines, tuple) and len(retainer_guidelines) > 1:
                            # Extract the ID (first element of the tuple)
                            parent_task_data["x_studio_guidelines"] = retainer_guidelines[0]
                        elif isinstance(retainer_guidelines, int):
                            parent_task_data["x_studio_guidelines"] = retainer_guidelines
                        else:
                            logger.warning(f"Guidelines not in expected format: {retainer_guidelines}")
                    
                    # Format dates correctly to avoid the microseconds issue
                    if retainer_request_receipt_dt:
                        parent_task_data["x_studio_request_receipt_date_time"] = retainer_request_receipt_dt.strftime("%Y-%m-%d %H:%M:%S")
                        
                    if retainer_internal_dt:
                        parent_task_data["x_studio_internal_due_date_1"] = retainer_internal_dt.strftime("%Y-%m-%d %H:%M:%S")
                    
                    # Create parent task in Odoo
                    parent_task_id = create_odoo_task(parent_task_data)
                    if not parent_task_id:
                        st.error("Failed to create parent task in Odoo.")
                        return
                        
                    st.success(f"Created Parent Task in Odoo (ID: {parent_task_id})")
                
                # STEP 2: CREATE SUBTASK
                with st.spinner("Creating subtask in Odoo..."):
                    # Create subtask with parent_id reference
                    subtask_data = {
                        "name": f"Retainer Subtask: {subtask_title}",
                        "project_id": project_id,
                        "parent_id": parent_task_id,  # This establishes the parent-child relationship
                        "user_ids": [(6, 0, [user_id])],
                        "description": f"Company: {selected_company}\nCustomer: {retainer_customer}\nProject: {parent_project_name}\nSubtask: {subtask_title}"
                    }
                    
                    # Add partner_id if found
                    if partner_id:
                        subtask_data["partner_id"] = partner_id
                    
                    # Add optional fields if they exist
                    if retainer_target_language:
                        subtask_data["x_studio_target_language"] = retainer_target_language
                    
                    # Handle guidelines properly - extract ID from tuple if applicable
                    if retainer_guidelines:
                        if isinstance(retainer_guidelines, tuple) and len(retainer_guidelines) > 1:
                            subtask_data["x_studio_guidelines"] = retainer_guidelines[0]
                        elif isinstance(retainer_guidelines, int):
                            subtask_data["x_studio_guidelines"] = retainer_guidelines
                    
                    # Format dates correctly
                    if retainer_request_receipt_dt:
                        subtask_data["x_studio_request_receipt_date_time"] = retainer_request_receipt_dt.strftime("%Y-%m-%d %H:%M:%S")
                        
                    if retainer_client_due_date_subtask:
                        subtask_data["x_studio_client_due_date_3"] = retainer_client_due_date_subtask.strftime("%Y-%m-%d")
                        
                    if retainer_internal_dt:
                        subtask_data["x_studio_internal_due_date_1"] = retainer_internal_dt.strftime("%Y-%m-%d %H:%M:%S")
                    
                    # Handle service category 1
                    if retainer_service_category_1:
                        if isinstance(retainer_service_category_1, tuple) and len(retainer_service_category_1) > 1:
                            if retainer_service_category_1[0] != -1:
                                subtask_data["x_studio_service_category_1"] = retainer_service_category_1[0]
                            else:
                                logger.warning(f"Skipping invalid service_category_1 ID: {retainer_service_category_1}")
                        elif isinstance(retainer_service_category_1, int):
                            subtask_data["x_studio_service_category_1"] = retainer_service_category_1
                        else:
                            logger.warning(f"Skipping service_category_1 as it's not in expected format: {retainer_service_category_1}")
                    
                    # Handle service category 2
                    if retainer_service_category_2:
                        if isinstance(retainer_service_category_2, tuple) and len(retainer_service_category_2) > 1:
                            if retainer_service_category_2[0] != -1:
                                subtask_data["x_studio_service_category_2"] = retainer_service_category_2[0]
                            else:
                                logger.warning(f"Skipping invalid service_category_2 ID: {retainer_service_category_2}")
                        elif isinstance(retainer_service_category_2, int):
                            subtask_data["x_studio_service_category_2"] = retainer_service_category_2
                        else:
                            logger.warning(f"Skipping service_category_2 as it's not in expected format: {retainer_service_category_2}")
                    
                    # Add design units
                    if no_of_design_units_sc1:
                        subtask_data["x_studio_total_no_of_design_units_sc1"] = no_of_design_units_sc1
                    
                    if no_of_design_units_sc2:
                        subtask_data["x_studio_total_no_of_design_units_sc2"] = no_of_design_units_sc2
                    
                    # Create subtask in Odoo
                    subtask_id = create_odoo_task(subtask_data)
                    if not subtask_id:
                        st.error("Failed to create subtask in Odoo.")
                        return
                        
                    st.success(f"Created Subtask in Odoo (ID: {subtask_id})")
                
                # STEP 3: CREATE GOOGLE DRIVE FOLDER STRUCTURE
                with st.spinner("Creating Google Drive folder structure for task..."):
                    # Sanitize folder name
                    folder_name = f"{parent_project_name} - {subtask_title} - {subtask_id}"
                    folder_name = folder_name.replace('/', '-').replace('\\', '-')
                    
                    # Create folder structure with subfolders
                    from google_drive import create_folder_structure
                    folder_structure = create_folder_structure(
                        folder_name, 
                        subfolders=["MATERIAL", "DELIVERABLE"]
                    )
                    
                    if folder_structure:
                        # Store main folder info in session state
                        st.session_state.drive_folder_id = folder_structure['main_folder_id']
                        st.session_state.drive_folder_link = folder_structure['main_folder_link']
                        st.session_state.folder_structure = folder_structure
                        # Update both parent and subtask with the folder links
                        try:
                            # Create a nicely formatted description with folder links
                            folder_description = f"\n\n📁 **Google Drive Folders:**\n"
                            folder_description += f"- Main Folder: {folder_structure['main_folder_url']}\n"
                            
                            # Add subfolder links if available
                            for subfolder_name, subfolder_info in folder_structure['subfolders'].items():
                                folder_description += f"- {subfolder_name}: {subfolder_info['url']}\n"
                            
                            # Update parent task description
                            parent_task_desc = models.execute_kw(
                                ODOO_DB, uid, ODOO_PASSWORD,
                                'project.task', 'read',
                                [[parent_task_id]],
                                {'fields': ['description']}
                            )[0]['description']
                            
                            updated_parent_desc = f"{parent_task_desc}{folder_description}"
                            
                            models.execute_kw(
                                ODOO_DB, uid, ODOO_PASSWORD,
                                'project.task', 'write',
                                [[parent_task_id], {'description': updated_parent_desc}]
                            )
                            
                            # Update subtask description
                            subtask_desc = models.execute_kw(
                                ODOO_DB, uid, ODOO_PASSWORD,
                                'project.task', 'read',
                                [[subtask_id]],
                                {'fields': ['description']}
                            )[0]['description']
                            
                            updated_subtask_desc = f"{subtask_desc}{folder_description}"
                            
                            models.execute_kw(
                                ODOO_DB, uid, ODOO_PASSWORD,
                                'project.task', 'write',
                                [[subtask_id], {'description': updated_subtask_desc}]
                            )
                            
                            logger.info(f"Updated tasks with Drive folder structure links")
                            st.success(f"Created folder structure with MATERIAL and DELIVERABLE subfolders")
                        except Exception as e:
                            logger.warning(f"Could not update tasks with folder links: {e}")
                    else:
                        st.warning("Could not create Google Drive folder. Please check logs for details.")
                
                # STEP 4: PREPARE FOR DESIGNER SELECTION
                with st.spinner("Preparing for designer selection..."):
                    # Fetch the task details from Odoo for designer selection
                    task_details = models.execute_kw(
                        ODOO_DB, uid, ODOO_PASSWORD,
                        'project.task', 'read',
                        [[subtask_id]],
                        {'fields': ['id', 'name', 'description', 'x_studio_service_category_1', 
                                'x_studio_service_category_2', 'x_studio_target_language',
                                'x_studio_client_due_date_3', 'date_deadline']}
                    )[0]
                    
                    # Store in session state
                    st.session_state.created_tasks = [task_details]
                    st.session_state.customer = retainer_customer
                    st.session_state.project = parent_project_name
                    st.session_state.parent_task_id = parent_task_id
                    st.session_state.designer_selection = True
                    
                    # Display designer suggestion if available
                    if "selected_designer" in st.session_state:
                        st.info(f"Suggested Designer: {st.session_state.selected_designer}")
                    
                    # Success message and transition
                    st.success("Tasks created successfully! Proceeding to designer selection...")
                    
                    st.rerun()
            
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
                logger.error(f"Error in retainer task creation: {e}", exc_info=True)
def inspect_field_values(models, uid, field_name, model_name='project.task', limit=50):
    """
    Inspects the values of a specific field across records to help diagnose type issues.
    
    Args:
        models: Odoo models proxy
        uid: User ID
        field_name: Name of the field to inspect
        model_name: Name of the model to inspect
        limit: Maximum number of records to fetch
        
    Returns:
        Analysis of the field values
    """
    try:
        # First, get field information
        field_info = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            model_name, 'fields_get',
            [[field_name]],
            {'attributes': ['string', 'type', 'relation', 'required', 'selection']}
        )
        
        if not field_info or field_name not in field_info:
            return f"Field '{field_name}' not found in model '{model_name}'."
            
        field_details = field_info[field_name]
        field_type = field_details.get('type', 'unknown')
        field_relation = field_details.get('relation', 'none')
        field_required = field_details.get('required', False)
        field_selection = field_details.get('selection', [])
        
        # Fetch records that have this field populated
        records = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            model_name, 'search_read',
            [[[(field_name, '!=', False)]]],
            {'fields': ['id', field_name, 'name'], 'limit': limit}
        )
        
        if not records:
            return (
                f"No records found with field '{field_name}' populated.\n"
                f"Field type: {field_type}\n"
                f"Relation model: {field_relation}\n"
                f"Required: {field_required}\n"
                f"Selection options: {field_selection}"
            )
            
        # Analyze the field values
        types_seen = {}
        sample_values = {}
        
        for rec in records:
            value = rec.get(field_name)
            value_type = type(value).__name__
            
            if value_type not in types_seen:
                types_seen[value_type] = 0
                sample_values[value_type] = []
                
            types_seen[value_type] += 1
            
            # Store a few sample values of each type
            if len(sample_values[value_type]) < 3:
                sample_values[value_type].append((rec.get('id'), rec.get('name', 'No Name'), value))
        
        # Generate report
        report = f"Field '{field_name}' Analysis (from {len(records)} records):\n\n"
        report += f"Field type in Odoo: {field_type}\n"
        report += f"Relation model: {field_relation}\n"
        report += f"Required: {field_required}\n"
        
        if field_selection:
            report += "Selection options:\n"
            for option in field_selection:
                report += f"  - {option}\n"
        
        report += "\nValue types found in database:\n"
        
        for t, count in types_seen.items():
            report += f"Type: {t} ({count} records)\n"
            report += "Sample values:\n"
            
            for record_id, record_name, sample in sample_values[t]:
                # Format the sample value for better readability
                if isinstance(sample, list):
                    sample_str = f"[{sample}]"
                else:
                    sample_str = str(sample)
                    
                report += f"  - ID {record_id} ({record_name}): {sample_str}\n"
                
            report += "\n"
            
        # If it's a relation field, provide some examples of how to use it
        if field_type in ['many2one', 'one2many', 'many2many']:
            report += f"\nThis is a {field_type} relation field. Here's how to use it:\n"
            
            if field_type == 'many2one':
                report += "For many2one fields, set the ID of the related record:\n"
                report += f"  task_data['{field_name}'] = 123  # ID of the related {field_relation} record\n"
            elif field_type == 'many2many':
                report += "For many2many fields, use a special command format:\n"
                report += f"  task_data['{field_name}'] = [(6, 0, [123, 456])]  # Replace with these IDs\n"
                report += f"  task_data['{field_name}'] = [(4, 123, 0)]  # Add this ID\n"
            elif field_type == 'one2many':
                report += "For one2many fields, use a special command format:\n"
                report += f"  task_data['{field_name}'] = [(0, 0, {{'field': 'value'}})]  # Create and link a new record\n"
                report += f"  task_data['{field_name}'] = [(1, 123, {{'field': 'value'}})]  # Update linked record with ID 123\n"
                
        return report
        
    except Exception as e:
        return f"Error inspecting field: {str(e)}"
    
def debug_task_fields():
    """
    Debug function to display all available fields on the project.task model
    """
    st.subheader("Debug: Task Fields")
    
    uid = st.session_state.odoo_uid
    models = st.session_state.odoo_models
    
    if not uid or not models:
        st.error("Not connected to Odoo. Please log in first.")
        return
    
    # Add a field inspection section
    st.subheader("Field Value Inspector")
    
    col1, col2 = st.columns(2)
    with col1:
        field_to_inspect = st.text_input("Enter field name to inspect:", value="x_studio_service_category_1")
    with col2:
        model_name = st.text_input("Model name:", value="project.task")
    
    if st.button("Inspect Field Values"):
        with st.spinner("Analyzing field values..."):
            report = inspect_field_values(models, uid, field_to_inspect, model_name)
            st.text_area("Field Analysis", report, height=400)
    
    # Add quick buttons for common problematic fields
    st.subheader("Quick Field Inspection")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("Inspect Service Category 1"):
            with st.spinner("Analyzing service category 1..."):
                report = inspect_field_values(models, uid, "x_studio_service_category_1")
                st.text_area("Service Category 1 Analysis", report, height=400)
    
    with col2:
        if st.button("Inspect Service Category 2"):
            with st.spinner("Analyzing service category 2..."):
                report = inspect_field_values(models, uid, "x_studio_service_category_2")
                st.text_area("Service Category 2 Analysis", report, height=400)
    
    with col3:
        if st.button("Inspect User IDs"):
            with st.spinner("Analyzing user_ids field..."):
                report = inspect_field_values(models, uid, "user_ids")
                st.text_area("User IDs Analysis", report, height=400)
                
    # Model discovery - helps find the actual models for related fields
    st.subheader("Model Discovery")
    
    model_prefix = st.text_input("Search for models with prefix:", value="x_")
    
    if st.button("Search Models"):
        with st.spinner("Searching models..."):
            try:
                # This gets all models in Odoo
                model_records = models.execute_kw(
                    ODOO_DB, uid, ODOO_PASSWORD,
                    'ir.model', 'search_read',
                    [[['model', 'like', model_prefix]]],
                    {'fields': ['id', 'model', 'name']}
                )
                
                if model_records:
                    st.success(f"Found {len(model_records)} models matching '{model_prefix}'")
                    
                    # Display in a table
                    model_data = []
                    for rec in model_records:
                        model_data.append({
                            "ID": rec.get('id'),
                            "Model": rec.get('model'),
                            "Name": rec.get('name')
                        })
                    
                    st.table(pd.DataFrame(model_data))
                else:
                    st.warning(f"No models found matching '{model_prefix}'")
            except Exception as e:
                st.error(f"Error searching models: {str(e)}")
    # Import needed constants
    from helpers import ODOO_DB, ODOO_PASSWORD
    
    try:
        with st.spinner("Fetching project.task fields..."):
            fields = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                'project.task', 'fields_get',
                [],
                {'attributes': ['string', 'type', 'required', 'relation']}
            )
        
        st.success(f"Found {len(fields)} fields on project.task")
        
        # Group fields by type for better display
        field_types = {}
        for field_name, field_info in fields.items():
            field_type = field_info.get('type', 'unknown')
            if field_type not in field_types:
                field_types[field_type] = []
            field_types[field_type].append((field_name, field_info))
        
        # Display fields by type
        for field_type, type_fields in field_types.items():
            with st.expander(f"{field_type.upper()} Fields ({len(type_fields)})"):
                for field_name, field_info in sorted(type_fields, key=lambda x: x[0]):
                    # Format the display of field information
                    field_label = field_info.get('string', 'No Label')
                    required = field_info.get('required', False)
                    relation = field_info.get('relation', 'None')
                    
                    # Highlight studio fields
                    if field_name.startswith('x_'):
                        st.markdown(f"**Field**: `{field_name}` - **Label**: {field_label}")
                        st.markdown(f"**Required**: {required} - **Relation**: {relation}")
                        st.markdown("---")
                    else:
                        st.markdown(f"Field: `{field_name}` - Label: {field_label}")
                        st.markdown(f"Required: {required} - Relation: {relation}")
                        st.markdown("---")
    
    except Exception as e:
        st.error(f"Error fetching field information: {str(e)}")
        logger.error(f"Error in debug_task_fields: {e}", exc_info=True)

def company_selection_page():
    st.title("Select Company")
    
    # Progress bar
    st.progress(25, text="Step 1 of 4: Select Company")

    # Connect to Odoo
    if "odoo_uid" not in st.session_state or "odoo_models" not in st.session_state:
        with st.spinner("Connecting to Odoo..."):
            uid, models = authenticate_odoo()
            if uid and models:
                st.session_state.odoo_uid = uid
                st.session_state.odoo_models = models
            else:
                st.error("Failed to connect to Odoo. Please check your credentials.")
                return
    else:
        uid = st.session_state.odoo_uid
        models = st.session_state.odoo_models

    # Fetch companies if not already in session
    if "companies" not in st.session_state:
        with st.spinner("Fetching companies..."):
            companies = get_companies(models, uid)
            if companies:
                st.session_state.companies = companies
            else:
                st.warning("No companies found or error fetching companies.")
                st.session_state.companies = []

    # Create form for company selection
    with st.form("company_selection_form"):
        st.subheader("Select Company")
        
        company_options = st.session_state.companies
        selected_company = st.selectbox(
            "Company",
            company_options
        )

        submit = st.form_submit_button("Next")
        
        if submit:
            # Validate selection
            if not selected_company:
                st.error("Please select a company.")
                return
                
            # Save to session state
            st.session_state.selected_company = selected_company
            st.session_state.company_selection_done = True
            
            st.success(f"Selected company: {selected_company}")
            st.rerun()

def email_analysis_page():
    st.title("Email Analysis")
    
    # Progress bar
    st.progress(66, text="Step 3 of 5: Email Analysis")
    
    # Navigation
    cols = st.columns([1, 5])
    with cols[0]:
        if st.button("← Back"):
            st.session_state["back_requested"] = True
            st.rerun()

    # Display current selection
    with st.container(border=True):
        selected_company = st.session_state.get("selected_company", "")
        parent_sales_order_item = st.session_state.get("parent_sales_order_item", "")
        customer = st.session_state.get("customer", "")
        project = st.session_state.get("project", "")
        
        st.markdown("### Current Selection")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f"**Company:** {selected_company}")
        with col2:
            st.markdown(f"**Sales Order:** {parent_sales_order_item}")
        with col3:
            st.markdown(f"**Customer:** {customer}")
        with col4:
            st.markdown(f"**Project:** {project}")
    
    # Suggested search terms OUTSIDE the form
    if parent_sales_order_item or customer:
        st.write("Suggested search terms:")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if parent_sales_order_item and st.button(f"Order: {parent_sales_order_item}", key="so_btn"):
                st.session_state.search_query = parent_sales_order_item
                st.rerun()
        with col2:
            if customer and st.button(f"Customer: {customer}", key="cust_btn"):
                st.session_state.search_query = customer
                st.rerun()
        with col3:
            if st.button("Recent emails", key="recent_btn"):
                st.session_state.search_query = ""
                st.rerun()
    
    # SKIP OPTION - always available outside any form
    if st.button("Skip Email Analysis", key="skip_outside"):
        st.session_state.email_analysis_done = True
        st.session_state.email_analysis_skipped = True
        st.rerun()
    
    # Check for Gmail authentication status BEFORE any forms
    gmail_authenticated = "google_gmail_creds" in st.session_state
    
    if not gmail_authenticated:
        # Authentication UI - completely outside of any form
        st.info("### Google Authentication Required")
        st.markdown("[Click here to authenticate with Google Gmail](https://prezlab-tms.streamlit.app/)")
        st.warning("You need to authenticate with Gmail before analyzing emails")
        return  # Exit the function early - don't show any forms
    
    # Only proceed with email fetching if already authenticated
    # Email Analysis Options Form
    with st.form("email_analysis_options"):
        st.subheader("Email Analysis Options")
        
        # Checkbox to enable/disable email analysis
        use_email_analysis = st.checkbox("Use Email Analysis to Enhance Task Details", value=True)
        
        # Group emails checkbox
        show_threads = st.checkbox("Group emails by conversation thread", value=True)
        
        if use_email_analysis:
            # Free text search - use session state value if set
            default_search = st.session_state.get("search_query", "")
            search_query = st.text_input("Search for emails (subject, sender, or content):",
                               value=default_search,
                               help="Type any keywords to find relevant emails")
            
            # Number of emails to fetch
            email_limit = st.slider("Number of emails to fetch", min_value=5, max_value=100, value=50)
        
        # Form submit button
        fetch_emails = st.form_submit_button("Fetch Emails")
    
    # Handle form submission OUTSIDE the form
    if fetch_emails and use_email_analysis:
        with st.spinner("Connecting to Gmail..."):
            try:
                # Gmail service initialization now happens outside any form
                gmail_service = get_gmail_service()
                
                if gmail_service:
                    with st.spinner(f"Fetching up to {email_limit} emails..."):
                        recent_emails = fetch_recent_emails(gmail_service, total_emails=email_limit, query=search_query)
                        
                        if show_threads:
                            from gmail_integration import extract_email_threads
                            threads = extract_email_threads(recent_emails)
                            st.session_state.email_threads = threads
                        
                        st.session_state.recent_emails = recent_emails
                        st.session_state.show_threads = show_threads
                        st.rerun()  # Refresh to show results
                else:
                    st.error("Failed to connect to Gmail. Please check your credentials.")
            except Exception as e:
                st.error(f"Error connecting to Gmail: {str(e)}")
    
    # Display emails or threads if fetched
    if "recent_emails" in st.session_state:
        if st.session_state.get("show_threads", False) and "email_threads" in st.session_state:
            # Display threads
            threads = st.session_state.email_threads
            thread_count = len(threads)
            
            if thread_count > 0:
                st.success(f"Found {thread_count} email threads")
                
                # Create a searchable dropdown
                thread_options = []
                thread_ids = list(threads.keys())
                
                for thread_id in thread_ids:
                    thread_emails = threads[thread_id]
                    # Use the most recent email in thread for display
                    latest_email = thread_emails[0] if thread_emails else {"from": "Unknown", "subject": "No Subject"}
                    option_text = f"{latest_email.get('from', 'Unknown')} - {latest_email.get('subject', 'No Subject')}"
                    thread_options.append(option_text)
                
                # Free text search for threads
                thread_search = st.text_input("Search within found threads:", 
                                     help="Type to filter the thread list below")
                
                # Filter the thread options based on search
                filtered_thread_options = []
                filtered_thread_indices = []
                
                for i, option in enumerate(thread_options):
                    if not thread_search or thread_search.lower() in option.lower():
                        filtered_thread_options.append(option)
                        filtered_thread_indices.append(i)
                
                if filtered_thread_options:
                    selected_option = st.selectbox(
                        "Select Thread to Analyze:",
                        options=range(len(filtered_thread_options)),
                        format_func=lambda i: filtered_thread_options[i]
                    )
                    
                    # Get the actual thread index
                    selected_thread_idx = filtered_thread_indices[selected_option]
                    selected_thread_id = thread_ids[selected_thread_idx]
                    selected_thread = threads[selected_thread_id]
                
                    # Show the thread content
                    # For threaded emails (replace the existing code)
                    with st.expander("View Thread Content", expanded=True):
                        for i, email in enumerate(selected_thread):
                            st.markdown(f"### Email {i+1}")
                            st.markdown(f"**From:** {email.get('from', 'Unknown')}")
                            st.markdown(f"**Subject:** {email.get('subject', 'No Subject')}")
                            st.markdown(f"**Date:** {email.get('date', 'Unknown')}")
                            
                            # Show email content
                            body = email.get('body', '')
                            if len(body) > 200:
                                st.markdown(f"**Content Preview:**\n{body[:200]}...")
                                # Replace nested expander with checkbox
                                show_full = st.checkbox(f"Show full content for Email {i+1}", key=f"show_full_thread_{i}")
                                if show_full:
                                    st.markdown(f"**Full Content:**\n{body}")
                            else:
                                st.markdown(f"**Content:**\n{body}")
                            
                            st.markdown("---")
                    
                    # Now create a form for analyzing - separate from thread selection
                    with st.form("analyze_thread_form"):
                        st.subheader("Analyze Selected Thread")
                        
                        # Email selection
                        st.write("Select emails to include in analysis:")
                        include_emails = []
                        for i, email in enumerate(selected_thread):
                            include = st.checkbox(f"Include Email {i+1}", value=True, key=f"email_{i}")
                            if include:
                                include_emails.append(i)
                        
                        # This is the submit button
                        analyze_thread_button = st.form_submit_button("Analyze Selected Emails")
                    
                    # Handle analysis OUTSIDE the form
                    if analyze_thread_button:
                        if include_emails:
                            with st.spinner("Analyzing emails with AI..."):
                                # Combine emails and analyze
                                combined_text = "Combined Email Thread:\n\n"
                                for i in include_emails:
                                    email = selected_thread[i]
                                    combined_text += f"Email {i+1}:\n"
                                    combined_text += f"From: {email.get('from', 'Unknown')}\n"
                                    combined_text += f"Subject: {email.get('subject', 'No Subject')}\n"
                                    combined_text += f"Content: {email.get('body', '')}\n\n"
                                
                                # Analyze the combined text
                                analysis_results = analyze_email(combined_text)
                                
                                # Store and proceed
                                st.session_state.email_analysis = analysis_results
                                st.session_state.email_analysis_done = True
                                st.session_state.email_analysis_skipped = False
                                st.rerun()
                        else:
                            st.warning("Please select at least one email to analyze.")
                else:
                    st.warning("No threads match your search. Try different terms.")
            else:
                st.warning("No email threads found. Try different search terms.")
                
                # Button to try again - outside any form
                if st.button("Try Different Search"):
                    st.session_state.pop("recent_emails", None)
                    st.session_state.pop("email_threads", None)
                    st.rerun()
        
        # Individual emails view (not threads)
        else:
            recent_emails = st.session_state.recent_emails
            if recent_emails:
                st.success(f"Found {len(recent_emails)} emails")
                
                # Create searchable dropdown for individual emails
                email_search = st.text_input("Search within found emails:", 
                                   help="Type to filter the email list below",
                                   key="email_search_indiv")
                
                # Filter emails based on search
                filtered_emails = []
                filtered_indices = []
                for i, email in enumerate(recent_emails):
                    subject = email.get('subject', 'No Subject')
                    sender = email.get('from', 'Unknown')
                    option_text = f"{sender} - {subject}"
                    
                    if not email_search or email_search.lower() in option_text.lower():
                        filtered_emails.append(option_text)
                        filtered_indices.append(i)
                
                if filtered_emails:
                    selected_email_option = st.selectbox(
                        "Select Email to Analyze:",
                        range(len(filtered_emails)),
                        format_func=lambda i: filtered_emails[i]
                    )
                    
                    # Get actual email index
                    selected_email_idx = filtered_indices[selected_email_option]
                    selected_email = recent_emails[selected_email_idx]
                    
                    # Show the selected email
                    # For individual emails (replace the existing code)
                    with st.expander("View Email Content", expanded=True):
                        st.markdown(f"**From:** {selected_email.get('from', 'Unknown')}")
                        st.markdown(f"**Subject:** {selected_email.get('subject', 'No Subject')}")
                        st.markdown(f"**Date:** {selected_email.get('date', 'Unknown')}")
                        
                        body = selected_email.get('body', '')
                        if len(body) > 200:
                            st.markdown(f"**Content Preview:**\n{body[:200]}...")
                            # Replace nested expander with checkbox
                            show_full = st.checkbox("Show full content", key=f"show_full_email_{selected_email_idx}")
                            if show_full:
                                st.markdown(f"**Full Content:**\n{body}")
                        else:
                            st.markdown(f"**Content:**\n{body}")
                    
                    # Single email analysis button - outside forms
                    if st.button("Analyze this Email", key="analyze_single_email"):
                        with st.spinner("Analyzing email with AI..."):
                            email_text = f"Subject: {selected_email.get('subject', '')}\n\nBody: {selected_email.get('body', '')}"
                            analysis_results = analyze_email(email_text)
                            
                            # Store analysis results
                            st.session_state.email_analysis = analysis_results
                            st.session_state.email_analysis_done = True
                            st.session_state.email_analysis_skipped = False
                            st.rerun()
                else:
                    st.warning("No emails match your search. Try different terms.")
            else:
                st.warning("No emails found. Try different search terms.")
                
                # Button to try again - outside any form
                if st.button("Try Different Search", key="try_diff_search_indiv"):
                    st.session_state.pop("recent_emails", None)
                    st.rerun()
    
    # Display analysis results if available
    if "email_analysis" in st.session_state and not st.session_state.get("email_analysis_skipped", True):
        analysis_results = st.session_state.email_analysis
        
        with st.container(border=True):
            st.subheader("Email Analysis Summary")
            
            if isinstance(analysis_results, dict):
                for key, value in analysis_results.items():
                    if value and key != "error" and key != "raw_analysis":
                        st.markdown(f"**{key.replace('_', ' ').title()}:** {value}")
            else:
                st.write(analysis_results)
        
        # Continue button - outside any form
        if st.button("Continue to Parent Task Details", type="primary"):
            st.rerun()

def google_auth_page():
    st.title("Google Services Authentication")
    
    # Check for both credential presence and backup flags
    gmail_authenticated = "google_gmail_creds" in st.session_state
    drive_authenticated = "google_drive_creds" in st.session_state
    
    # Check for backup auth completion flags
    gmail_auth_complete = st.session_state.get("gmail_auth_complete", False)
    drive_auth_complete = st.session_state.get("drive_auth_complete", False)
    
    # Add debug info to help troubleshoot
    with st.sidebar.expander("Debug Info", expanded=False):
        st.write("Session State Keys:", list(st.session_state.keys()))
        st.write("Gmail Creds:", gmail_authenticated)
        st.write("Drive Creds:", drive_authenticated)
        st.write("Gmail Auth Complete Flag:", gmail_auth_complete)
        st.write("Drive Auth Complete Flag:", drive_auth_complete)
        
        # Add a reset button
        if st.button("Reset Google Auth Status (Debug)"):
            keys_to_remove = [
                'google_auth_complete', 
                'google_gmail_creds', 
                'google_drive_creds',
                'gmail_auth_complete', 
                'drive_auth_complete'
            ]
            for key in keys_to_remove:
                if key in st.session_state:
                    st.session_state.pop(key, None)
            st.rerun()
    
    st.info("Please authenticate with Google to enable Gmail and Drive functionality.")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Gmail")
        if gmail_authenticated or gmail_auth_complete:
            st.success("✅ Authenticated")
        else:
            st.warning("⚠️ Not authenticated")
            if st.button("Authenticate Gmail"):
                with st.spinner("Connecting to Gmail..."):
                    gmail_service = get_gmail_service()
                    if gmail_service:
                        # Set both the credential and backup flags
                        st.session_state.gmail_auth_complete = True
                        st.success("Gmail authentication successful!")
                        st.rerun()
    
    with col2:
        st.subheader("Google Drive")
        if drive_authenticated or drive_auth_complete:
            st.success("✅ Authenticated")
        else:
            st.warning("⚠️ Not authenticated")
            if st.button("Authenticate Drive"):
                with st.spinner("Connecting to Drive..."):
                    drive_service = get_drive_service()
                    if drive_service:
                        # Set both the credential and backup flags
                        st.session_state.drive_auth_complete = True
                        st.success("Drive authentication successful!")
                        st.rerun()
    
    # Check if both services are authenticated
    if (gmail_authenticated or gmail_auth_complete) and (drive_authenticated or drive_auth_complete):
        st.success("All services authenticated! You're ready to proceed.")
        
        # Set the main flag that the main() function checks
        st.session_state.google_auth_complete = True
        
        if st.button("Continue to Dashboard", type="primary"):
            # Ensure the flag is set before continuing
            st.session_state.google_auth_complete = True
            st.rerun()
    else:
        # Make sure google_auth_complete stays False until both services are authenticated
        st.session_state.google_auth_complete = False

def initialize_gmail_connection():
    """
    Initialize connection to Gmail API and handle authentication.
    Returns the Gmail service object or None if connection fails.
    """
    try:
        service = get_gmail_service()
        if service:
            return service
        else:
            st.error("Failed to initialize Gmail service. Check your credentials.")
            return None
    except Exception as e:
        st.error(f"Error connecting to Gmail: {str(e)}")
        return None
    
def designer_selection_page():
    """
    Page for selecting and booking designers for created tasks.
    Displayed after task creation in both ad-hoc and retainer flows.
    """
    st.title("Designer Selection & Booking")

    # Check if we have tasks to assign designers to
    if "created_tasks" not in st.session_state or not st.session_state.created_tasks:
        st.warning("No tasks available for designer assignment. Please create tasks first.")
        # Clear the designer_selection flag to prevent coming back here
        if "designer_selection" in st.session_state:
            st.session_state.pop("designer_selection", None)
            
        if st.button("Return to Home"):
            for key in ["form_type", "adhoc_sales_order_done", "adhoc_parent_input_done", 
                      "retainer_parent_input_done", "subtask_index", "created_tasks", 
                      "company_selection_done", "email_analysis_done", "email_analysis_skipped"]:
                st.session_state.pop(key, None)
            st.rerun()
        return
    
    # Get Odoo connection
    uid = st.session_state.odoo_uid
    models = st.session_state.odoo_models
    
    # Display created tasks
    st.subheader("Tasks Ready for Designer Assignment")
    
    tasks = st.session_state.created_tasks
    parent_task_id = st.session_state.get("parent_task_id")
    
    # Display parent task info if available
    if parent_task_id:
        try:
            # Fetch parent task details
            parent_task_data = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                'project.task', 'read',
                [[parent_task_id]],
                {'fields': ['name', 'x_studio_service_category_1', 'x_studio_service_category_2', 'description']}
            )
            
            if parent_task_data:
                parent_info = parent_task_data[0]
                with st.container(border=True):
                    st.markdown(f"**Parent Task:** {parent_info.get('name', f'ID: {parent_task_id}')}")
                    st.markdown(f"**Parent Task ID:** {parent_task_id}")
                    st.markdown(f"**Project:** {st.session_state.get('project', '')}")
                    st.markdown(f"**Customer:** {st.session_state.get('customer', '')}")
                    
                    # Service categories if available
                    if 'x_studio_service_category_1' in parent_info and parent_info['x_studio_service_category_1']:
                        if isinstance(parent_info['x_studio_service_category_1'], list):
                            st.markdown(f"**Service Category 1:** {parent_info['x_studio_service_category_1'][1]}")
                        else:
                            st.markdown(f"**Service Category 1:** {parent_info['x_studio_service_category_1']}")
                    
                    # Add Drive folder link if available
                    if "drive_folder_link" in st.session_state:
                        st.markdown(f"**📁 Google Drive Folder:** [Open Folder]({st.session_state.drive_folder_link})")
        except Exception as e:
            logger.warning(f"Could not fetch parent task details: {e}")
            with st.container(border=True):
                st.markdown(f"**Parent Task ID:** {parent_task_id}")
                st.markdown(f"**Project:** {st.session_state.get('project', '')}")
                st.markdown(f"**Customer:** {st.session_state.get('customer', '')}")
                if "drive_folder_link" in st.session_state:
                    st.markdown(f"**📁 Google Drive Folder:** [Open Folder]({st.session_state.drive_folder_link})")

    # Load all designers once
    with st.spinner("Loading designer information..."):
        try:
            designers_df = load_designers()
            if designers_df.empty:
                st.error("No designer information available. Please check the designer data file.")
                return
            
            # Get all employees from planning for availability check
            employees = get_all_employees_in_planning(models, uid)
        except Exception as e:
            st.error(f"Error loading designers: {str(e)}")
            logger.error(f"Error loading designers: {e}", exc_info=True)
            return
    
    # Iterate through each task to assign designers
    for i, task in enumerate(tasks):
        st.markdown(f"### Task {i+1}: {task.get('name', f'Task ID: {task.get('id', 'Unknown')}')}")
        
        # Get task details in a format suitable for designer matching
        task_details = f"Task: {task.get('name', '')}\n"
        task_details += f"Description: {task.get('description', '')}\n"
        
        # Add additional details if available
        for field, label in [
            ('x_studio_service_category_1', 'Service Category 1'),
            ('x_studio_service_category_2', 'Service Category 2'),
            ('x_studio_target_language', 'Target Language')
        ]:
            if field in task and task[field]:
                if isinstance(task[field], list) and len(task[field]) >= 2:
                    task_details += f"{label}: {task[field][1]}\n"
                else:
                    task_details += f"{label}: {task[field]}\n"
        
        # Add parent task information to task details for better matching
        if parent_task_id:
            task_details += f"Parent Task ID: {parent_task_id}\n"
            task_details += f"Project: {st.session_state.get('project', '')}\n"
            task_details += f"Customer: {st.session_state.get('customer', '')}\n"
        
        # Show current task status
        current_status = "Not assigned"
        if "designer_assigned" in task and task["designer_assigned"]:
            current_status = f"Assigned to {task['designer_assigned']}"
        
        st.markdown(f"**Status:** {current_status}")
        
        # Designer suggestion and assignment section
        with st.container(border=True):
            col1, col2 = st.columns([2, 1])
            
            with col1:
                # Only suggest if not already assigned
                if "designer_assigned" not in task or not task["designer_assigned"]:
                    if st.button(f"Suggest Designers for Task {i+1}", key=f"suggest_{task['id']}"):
                        with st.spinner("Analyzing best designers..."):
                            # Calculate due date for availability check
                            due_date = None
                            if 'x_studio_client_due_date_3' in task and task['x_studio_client_due_date_3']:
                                due_date = task['x_studio_client_due_date_3']
                            elif 'date_deadline' in task and task['date_deadline']:
                                due_date = task['date_deadline']
                            else:
                                due_date = datetime.now() + pd.Timedelta(days=7)  # Default 1 week
                            
                            # Estimate task duration based on service categories and design units
                            estimated_duration = 8  # Default 8 hours
                            
                            # Try to refine duration estimate based on design units if available
                            design_units_sc1 = task.get('x_studio_total_no_of_design_units_sc1', 0)
                            design_units_sc2 = task.get('x_studio_total_no_of_design_units_sc2', 0)
                            
                            if design_units_sc1 or design_units_sc2:
                                # Simple formula: 2 hours per design unit, minimum 4 hours
                                total_units = (design_units_sc1 or 0) + (design_units_sc2 or 0)
                                if total_units > 0:
                                    estimated_duration = max(4, total_units * 2)
                            
                            # Get ranked designers
                            ranked_designers = rank_designers_by_skill_match(task_details, designers_df)
                            
                            # Filter by availability
                            available_designers, unavailable_designers = filter_designers_by_availability(
                                ranked_designers, models, uid, due_date, estimated_duration
                            )
                            
                            # Store in session state for this task
                            task_key = f"designer_options_{task['id']}"
                            st.session_state[task_key] = {
                                'available': available_designers,
                                'unavailable': unavailable_designers,
                                'task_details': task_details,
                                'due_date': due_date,
                                'duration': estimated_duration
                            }
                            
                            # Get the full suggestion with availability info
                            suggestion = suggest_best_designer_available(
                                task_details, available_designers, unavailable_designers
                            )
                            
                            # Store the suggestion
                            st.session_state[f"designer_suggestion_{task['id']}"] = suggestion
                            
                            # Refresh to show results
                            st.rerun()
            
            with col2:
                # Show scheduling button if not already assigned
                if "designer_assigned" not in task or not task["designer_assigned"]:
                    if st.button(f"Schedule Task {i+1}", key=f"schedule_{task['id']}"):
                        # This will be handled below after displaying all options
                        st.session_state[f"schedule_task_{task['id']}"] = True
                        st.rerun()
        
        # Display designer suggestions if available
        designer_key = f"designer_options_{task['id']}"
        suggestion_key = f"designer_suggestion_{task['id']}"
        
        if suggestion_key in st.session_state:
            st.markdown("#### Designer Recommendation")
            st.info(st.session_state[suggestion_key])
        
        # Display designer selection options if available
        if designer_key in st.session_state:
            options = st.session_state[designer_key]
            
            available_df = options['available']
            
            if not available_df.empty:
                st.markdown("#### Available Designers")
                
                # Display a selection widget for available designers
                designer_options = []
                for _, row in available_df.iterrows():
                    name = row['Name']
                    score = row.get('match_score', 0)
                    avail = row.get('available_from', '')
                    
                    if avail and isinstance(avail, pd.Timestamp):
                        avail_str = avail.strftime("%Y-%m-%d %H:%M")
                    else:
                        avail_str = "Unknown"
                        
                    option_text = f"{name} (Match: {score:.0f}%, Available: {avail_str})"
                    designer_options.append((name, option_text))
                
                # Selectbox for designer selection
                if designer_options:
                    selected_option = st.selectbox(
                        "Select Designer:",
                        options=range(len(designer_options)),
                        format_func=lambda i: designer_options[i][1],
                        key=f"designer_select_{task['id']}"
                    )
                    
                    # Store the selected designer name
                    selected_designer = designer_options[selected_option][0]
                    st.session_state[f"selected_designer_{task['id']}"] = selected_designer
            else:
                st.warning("No designers are currently available for this task.")
        
        # Handle scheduling if requested
        if f"schedule_task_{task['id']}" in st.session_state and st.session_state[f"schedule_task_{task['id']}"]:
            selected_designer_key = f"selected_designer_{task['id']}"
            
            if selected_designer_key in st.session_state:
                designer_name = st.session_state[selected_designer_key]
                
                # Add debug logging to verify the correct designer name is being passed
                st.write(f"Debug: Selected designer: {designer_name}")
                
                # Find employee ID with more robust matching
                employee_id = None
                for emp in employees:
                    if designer_name.lower() in emp['name'].lower():
                        employee_id = emp['id']
                        st.write(f"Debug: Found employee ID: {employee_id} for {emp['name']}")
                        break

                # Try to find the available slot for this designer
                designer_key = f"designer_options_{task['id']}"
                if designer_key in st.session_state:
                    options = st.session_state[designer_key]
                    available_df = options['available']
                    
                    # Find the designer's row
                    designer_row = available_df[available_df['Name'] == designer_name]
                    
                    if not designer_row.empty:
                        # Get availability information
                        available_from = designer_row.iloc[0].get('available_from')
                        available_until = designer_row.iloc[0].get('available_until')
                        
                        # Schedule the task
                        with st.spinner(f"Scheduling task for {designer_name}..."):
                            # Find employee ID
                            employee_id = find_employee_id(designer_name, employees)
                            
                            if employee_id:
                                # Get parent task name if available
                                parent_info = ""
                                if parent_task_id:
                                    try:
                                        parent_task_data = models.execute_kw(
                                            ODOO_DB, uid, ODOO_PASSWORD,
                                            'project.task', 'read',
                                            [[parent_task_id]],
                                            {'fields': ['name']}
                                        )
                                        if parent_task_data:
                                            parent_name = parent_task_data[0].get('name', f"Parent {parent_task_id}")
                                            parent_info = f" | Parent: {parent_name}"
                                    except Exception as e:
                                        logger.warning(f"Could not fetch parent task name: {e}")
                                        parent_info = f" | Parent ID: {parent_task_id}"
                                
                                # Get task name with project/customer context
                                task_name = task.get('name', f"Task {task['id']}")
                                project_name = st.session_state.get('project', '')
                                customer_name = st.session_state.get('customer', '')
                                
                                planning_task_name = f"{task_name}{parent_info}"
                                if project_name:
                                    planning_task_name += f" | Project: {project_name}"
                                if customer_name:
                                    planning_task_name += f" | Customer: {customer_name}"
                                
                                # Convert to datetime if needed
                                task_start = available_from
                                task_end = available_until
                                
                                if isinstance(task_start, pd.Timestamp):
                                    task_start = task_start.to_pydatetime()
                                if isinstance(task_end, pd.Timestamp):
                                    task_end = task_end.to_pydatetime()
                                

                                # Add before creating the planning slot
                                planning_fields = get_available_fields(models, uid, 'planning.slot')
                                st.write("Available planning.slot fields:", list(planning_fields.keys()))

                                # Create the planning slot with full context
                                slot_id = create_task(
                                    models, uid, employee_id, 
                                    planning_task_name, 
                                    task_start, task_end,
                                    parent_task_id=parent_task_id,
                                    task_id=task['id']
                                )
                                
                                if slot_id:
                                    # Update the task with the assigned designer
                                    with st.spinner(f"Updating task with designer {designer_name}..."):
                                        # Create a more descriptive assignment note
                                        assignment_note = (
                                            f"Designer {designer_name} assigned by Task Management System\n"
                                            f"Assignment Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
                                            f"Scheduled Time: {task_start.strftime('%Y-%m-%d %H:%M')} to {task_end.strftime('%Y-%m-%d %H:%M')}\n"
                                            f"Planning Slot ID: {slot_id}"
                                        )
                                        
                                        success = update_task_designer(models, uid, task['id'], designer_name, 
                                                                      assignment_note=assignment_note, 
                                                                      planning_slot_id=slot_id,
                                                                      role="Designer")
                                        
                                        if success:
                                            st.success(f"Successfully assigned {designer_name} to the task!")
                                            # Update task in session state
                                            task["designer_assigned"] = designer_name
                                            task["planning_slot_id"] = slot_id
                                            # Clean up session state keys for this task
                                            st.session_state.pop(f"schedule_task_{task['id']}", None)
                                        else:
                                            st.error(f"Failed to update task with designer information. Please check Odoo configuration and field names.")
                                            # Provide troubleshooting information
                                            with st.expander("Troubleshooting Information"):
                                                st.markdown("The system failed to update the task with the designer information. This could be due to:")
                                                st.markdown("1. Missing fields in the Odoo task model")
                                                st.markdown("2. Insufficient permissions for the current Odoo user")
                                                st.markdown("3. Network connectivity issues")
                                                st.markdown("4. The designer may not have a corresponding user record in Odoo")
                                                st.markdown("\nCheck the application logs for more detailed error information.")
                                else:
                                    st.error(f"Failed to create planning slot for {designer_name}.")
                            else:
                                st.error(f"Could not find {designer_name} in the planning system.")
                    else:
                        st.error(f"Designer {designer_name} is no longer available.")
            else:
                st.warning("Please select a designer first.")
                st.session_state.pop(f"schedule_task_{task['id']}", None)
        
        st.markdown("---")
    
    # Final success and navigation section
    if all("designer_assigned" in task and task["designer_assigned"] for task in tasks):
        st.success("All tasks have been assigned to designers!")
        
        if st.button("Complete Process", type="primary"):
            # Clear ALL relevant session state keys
            keys_to_clear = [
                "form_type", 
                "adhoc_sales_order_done", 
                "adhoc_parent_input_done",
                "retainer_parent_input_done", 
                "subtask_index", 
                "created_tasks",
                "designer_selection",  # Make sure this flag is cleared
                "parent_task_id",
                "company_selection_done",
                "email_analysis_done",
                "email_analysis_skipped"
            ]
            
            for key in keys_to_clear:
                st.session_state.pop(key, None)
            st.rerun()
    else:
        # If not all tasks are assigned, show appropriate message
        assigned_count = sum(1 for task in tasks if "designer_assigned" in task and task["designer_assigned"])
        st.info(f"{assigned_count} out of {len(tasks)} tasks have been assigned to designers.")
        
        if st.button("Complete Process (Skip Remaining Assignments)", type="primary"):
            # Clear session state and return to home even if not all tasks are assigned
            for key in ["form_type", "adhoc_sales_order_done", "adhoc_parent_input_done", 
                      "retainer_parent_input_done", "subtask_index", "created_tasks",
                      "designer_selection", "parent_task_id"]:
                st.session_state.pop(key, None)
            st.rerun()


# -------------------------------
# MAIN
# -------------------------------

def main():
    """
    Entry‑point for the Streamlit Task‑Management app.
    ‑   Ensures the user is logged‑in *before* we complete any Google OAuth
    ‑   Persists OAuth codes that arrive early, then processes them post‑login
    """
    from session_manager import SessionManager
    SessionManager.initialize_session()

    import streamlit as st

    inject_custom_css()


    # Set logo position based on login status
    if st.session_state.get("logged_in", False):
        # For logged-in pages, use top-right position
        add_logo(position="top-right", width=120)
    else:
        # For login page, this will be handled in login_page()
        pass
    # ------------------------------------------------------------------
    # 1)  Capture *early* Google OAuth callback codes
    # ------------------------------------------------------------------
    if "code" in st.query_params:
        # Stash it until the user finishes the TMS login
        st.session_state.pending_oauth_code = st.query_params["code"]
        st.query_params.clear()               # keeps the URL tidy

    # ------------------------------------------------------------------
    # 2)  Admin “system debug” shortcut (unchanged)
    # ------------------------------------------------------------------
    if inject_debug_page():                   # noqa: F821  (defined earlier in app.py)
        return

    # ------------------------------------------------------------------
    # 3)  Sidebar is always visible
    # ------------------------------------------------------------------
    render_sidebar()                          # noqa: F821

    # Small live‑state panel
    st.sidebar.write("Debug Info:")
    st.sidebar.write(f"Logged in: {st.session_state.get('logged_in', False)}")
    st.sidebar.write(f"Google Auth Complete: {st.session_state.get('google_auth_complete', False)}")
    st.sidebar.write(f"Google Gmail Creds: {'google_gmail_creds' in st.session_state}")
    st.sidebar.write(f"Google Drive Creds: {'google_drive_creds' in st.session_state}")

    # ------------------------------------------------------------------
    # 4)  Login gate
    # ------------------------------------------------------------------
    if not st.session_state.get("logged_in", False):
        login_page()                          # noqa: F821
        return

    # ------------------------------------------------------------------
    # 5)  *After* successful login, finish any pending OAuth handshake
    # ------------------------------------------------------------------
    if "pending_oauth_code" in st.session_state:
        code = st.session_state.pop("pending_oauth_code")

        from google_auth import handle_oauth_callback  # local import avoids circular refs
        st.info("Finishing Google authentication…")
        success = handle_oauth_callback(code)

        if success:
            st.success("Google account linked – token saved to Supabase.")
        else:
            st.error("Google authentication failed. Please try again.")

        st.rerun()     # refresh state / UI regardless of outcome
        return

    # ------------------------------------------------------------------
    # 6)  Optional Auth‑debug / Google‑auth pages (unchanged)
    # ------------------------------------------------------------------
    if st.session_state.get("show_google_auth", False):
        google_auth_page()                    # noqa: F821
        if st.button("Return to Main Page"):
            st.session_state.pop("show_google_auth", None)
            st.rerun()
        return
    elif st.session_state.get("debug_mode") == "auth_debug":
        auth_debug_page()                     # noqa: F821

    # Additional debug mode (task‑fields) — unchanged
    if st.session_state.get("debug_mode") == "task_fields":
        debug_task_fields()                   # noqa: F821
        if st.button("Return to Normal Mode"):
            st.session_state.pop("debug_mode")
            st.rerun()
        return

    # ------------------------------------------------------------------
    # 7)  Session‑validation & main workflow (unchanged)
    # ------------------------------------------------------------------
    if not validate_session():                # noqa: F821
        login_page()
        return
    

    # Main content routing (identical to previous logic)
    if "form_type" not in st.session_state:
        type_selection_page()                 # noqa: F821
    else:
        # Designer‑selection page shown after tasks are created
        if st.session_state.get("designer_selection"):
            designer_selection_page()         # noqa: F821
        elif st.session_state.form_type == "Via Sales Order":
            if "adhoc_parent_input_done" in st.session_state:
                adhoc_subtask_page()
            elif "adhoc_sales_order_done" in st.session_state:
                if "email_analysis_done" in st.session_state:
                    adhoc_parent_task_page()
                else:
                    email_analysis_page()
            elif "company_selection_done" in st.session_state:
                sales_order_page()
            else:
                company_selection_page()
        elif st.session_state.form_type == "Via Project":
            if "retainer_parent_input_done" in st.session_state:
                retainer_subtask_page()       # noqa: F821
            else:
                retainer_parent_task_page()   # noqa: F821

if __name__ == "__main__":
    main()