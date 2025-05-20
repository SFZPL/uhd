import os
import xmlrpc.client
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import logging
from typing import Dict, List, Tuple, Optional, Any, Union

# ────────────────────────────────────────────────────────────
# Only load .env in local dev, skip on Streamlit Cloud
try:
    if os.getenv("LOCAL_DEVELOPMENT", "False").lower() in ("true", "1"):
        from dotenv import load_dotenv
        load_dotenv()
        logging.getLogger(__name__).info("Loaded .env for local development")
except ImportError:
    # python‐dotenv not installed in production — ignore
    pass
except Exception as e:
    logging.getLogger(__name__).warning(f"Skipping load_dotenv(): {e}")
# ────────────────────────────────────────────────────────────

from config import get_secret

# Initialize Odoo credential globals (populated at runtime)
ODOO_URL = ODOO_DB = ODOO_USERNAME = ODOO_PASSWORD = None

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    filename="helpers.log"
)
logger = logging.getLogger(__name__)

# Type definitions
OdooConnection = Tuple[int, xmlrpc.client.ServerProxy]
OdooRecord     = Dict[str, Any]


# helpers.py
# ------------------------------------------------------------
# Odoo connection helper
# ------------------------------------------------------------
# helpers.py
# ------------------------------------------------------------
# Odoo connection helper (re‑worked)
# ------------------------------------------------------------
def get_odoo_connection(force_refresh: bool = False):
    """
    Return (uid, models) if successful, otherwise (None, None).
    Pulls all four Odoo secrets at call time and updates the module globals.
    Caches the connection in st.session_state for up to 1 hour.
    """
    global ODOO_URL, ODOO_DB, ODOO_USERNAME, ODOO_PASSWORD

    # ─── 1) Fetch secrets afresh ────────────────────────────────────
    ODOO_URL      = get_secret("ODOO_URL")
    ODOO_DB       = get_secret("ODOO_DB")
    ODOO_USERNAME = get_secret("ODOO_USERNAME")
    ODOO_PASSWORD = get_secret("ODOO_PASSWORD")

    # Fail fast if any are missing
    missing = [k for k,v in [
        ("ODOO_URL",      ODOO_URL),
        ("ODOO_DB",       ODOO_DB),
        ("ODOO_USERNAME", ODOO_USERNAME),
        ("ODOO_PASSWORD", ODOO_PASSWORD)
    ] if not v]
    if missing:
        logger.error(f"Missing Odoo secrets: {', '.join(missing)}")
        return None, None

    # ─── 2) Return cached if still fresh ───────────────────────────
    if not force_refresh and "odoo_connection" in st.session_state:
        conn = st.session_state.odoo_connection
        if (datetime.now() - conn["timestamp"]) < timedelta(hours=1):
            return conn["uid"], conn["models"]

    # ─── 3) Establish a new XML‑RPC connection ─────────────────────
    try:
        logger.info("Establishing new Odoo XML‑RPC connection")
        common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common", allow_none=True)

        uid = common.authenticate(ODOO_DB, ODOO_USERNAME, ODOO_PASSWORD, {})
        if not uid:
            raise RuntimeError("authenticate() returned False")

        models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object", allow_none=True)

        # Cache it in Streamlit session state
        st.session_state.odoo_connection = {
            "uid":       uid,
            "models":    models,
            "timestamp": datetime.now(),
        }
        logger.info(f"Odoo connection successful (UID {uid})")
        return uid, models

    except xmlrpc.client.ProtocolError as e:
        logger.error(f"Odoo protocol error {e.errcode}: {e.errmsg}", exc_info=True)
        return None, None
    except Exception as e:
        logger.error(f"Odoo connection error: {e}", exc_info=True)
        return None, None


def check_odoo_connection():
    """
    Validates that the Odoo connection is active
    
    Returns:
        True if connection is valid, False otherwise
    """
    uid, models = get_odoo_connection()
    if not uid or not models:
        return False
        
    try:
        # Simple test query to validate connection
        result = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'res.users', 'search_count',
            [[['id', '=', uid]]]
        )
        return result == 1
    except Exception:
        # If any error occurs, connection is invalid
        return False

def authenticate_odoo() -> OdooConnection:
    """Alias for get_odoo_connection for backward compatibility"""
    return get_odoo_connection()

def create_odoo_task(task_data: Dict[str, Any]) -> Optional[int]:
    """
    Creates a task in Odoo with the provided data.
    
    Args:
        task_data: Dictionary containing task field values
        
    Returns:
        Task ID if successful, None if failed
    """
    uid, models = get_odoo_connection()
    if uid is None or models is None:
        st.error("Odoo connection failed.")
        return None
    
    # Log task data for debugging
    logger.info(f"Creating task with data: {task_data}")
    
    try:
        # Sanitize data - ensure all fields have appropriate types
        sanitized_data = {}
        for key, value in task_data.items():
            if key == "project_id" and not isinstance(value, int):
                try:
                    sanitized_data[key] = int(value)
                except (ValueError, TypeError):
                    logger.error(f"Invalid project_id: {value}")
                    st.error(f"Invalid project ID: {value}")
                    return None
            elif key == "user_ids" and isinstance(value, list):
                # Ensure user_ids has correct format for many2many fields
                if value and isinstance(value[0], tuple) and len(value[0]) >= 3:
                    sanitized_data[key] = value
                else:
                    logger.error(f"Invalid user_ids format: {value}")
                    st.error(f"Invalid user IDs format: {value}")
                    return None
            else:
                sanitized_data[key] = value
        
        # Create task in Odoo
        task_id = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'project.task', 'create', [sanitized_data]
        )
        
        logger.info(f"Task created successfully (ID: {task_id})")
        return task_id
        
    except Exception as e:
        logger.error(f"Error creating task in Odoo: {e}", exc_info=True)
        st.error(f"Error creating task in Odoo: {e}")
        return None

# Modify get_sales_orders in helpers.py to filter by company
def get_sales_orders(models: xmlrpc.client.ServerProxy, uid: int, company_name: str = None) -> List[OdooRecord]:
    """
    Retrieves a list of sales orders from Odoo, optionally filtered by company.
    
    Args:
        models: Odoo models proxy
        uid: User ID
        company_name: Optional company name to filter by
        
    Returns:
        List of sales order records
    """
    try:
        # If company_name is provided, add it to the domain filter
        domain = []
        if company_name:
            domain = [('company_id.name', '=', company_name)]
            
        orders = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'sale.order', 'search_read',
            [domain],
            {'fields': ['name', 'partner_id', 'project_id']}
        )
        logger.info(f"Retrieved {len(orders)} sales orders")
        return orders
        
    except Exception as e:
        logger.error(f"Error fetching sales orders: {e}", exc_info=True)
        st.warning("Error fetching sales orders, reauthenticating...")
        
        # Retry with new connection
        uid, models = authenticate_odoo()
        try:
            domain = []
            if company_name:
                domain = [('company_id.name', '=', company_name)]
                
            orders = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                'sale.order', 'search_read',
                [domain],
                {'fields': ['name', 'partner_id', 'project_id']}
            )
            return orders
        except Exception as e:
            logger.error(f"Retry failed to fetch sales orders: {e}", exc_info=True)
            st.error(f"Failed to retrieve sales orders: {e}")
            return []
        
def get_sales_order_details(models: xmlrpc.client.ServerProxy, uid: int, sales_order_name: str) -> Dict[str, str]:
    """
    Retrieves details for a specific sales order.
    
    Args:
        models: Odoo models proxy
        uid: User ID
        sales_order_name: Name of the sales order
        
    Returns:
        Dictionary containing sales order details
    """
    try:
        orders = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'sale.order', 'search_read',
            [[['name', '=', sales_order_name]]],
            {'fields': ['name', 'partner_id', 'project_id']}
        )
        
        if orders:
            order = orders[0]
            details = {}
            details['sales_order'] = order.get('name', '')
            
            # Extract partner name
            partner = order.get('partner_id', [0, ''])
            details['customer'] = partner[1] if isinstance(partner, list) and len(partner) > 1 else ""
            
            # Extract project name
            project = order.get('project_id', [0, ''])
            details['project'] = project[1] if isinstance(project, list) and len(project) > 1 else ""
            
            logger.info(f"Retrieved details for sales order: {sales_order_name}")
            return details
        return {}
        
    except Exception as e:
        logger.error(f"Error fetching sales order details: {e}", exc_info=True)
        st.error(f"Error retrieving sales order details: {e}")
        return {}

def get_employee_schedule(models: xmlrpc.client.ServerProxy, uid: int, employee_id: int) -> List[OdooRecord]:
    """
    Retrieves the schedule for a specific employee.
    
    Args:
        models: Odoo models proxy
        uid: User ID
        employee_id: Employee ID
        
    Returns:
        List of schedule records
    """
    try:
        tasks = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'planning.slot', 'search_read',
            [[['resource_id', '=', employee_id]]],
            {'fields': ['start_datetime', 'end_datetime'], 'order': 'start_datetime'}
        )
        logger.info(f"Retrieved {len(tasks)} scheduled tasks for employee {employee_id}")
        return tasks
        
    except Exception as e:
        logger.error(f"Error fetching employee schedule: {e}", exc_info=True)
        st.error(f"Error retrieving employee schedule: {e}")
        return []

def create_task(models: xmlrpc.client.ServerProxy, uid: int, employee_id: int, 
                task_name: str, task_start: datetime, task_end: datetime, 
                parent_task_id: int = None, task_id: int = None) -> Optional[int]:
    """
    Creates a new task in the employee's schedule.
    
    Args:
        models: Odoo models proxy
        uid: User ID
        employee_id: Employee ID
        task_name: Name of the task
        task_start: Start datetime
        task_end: End datetime
        parent_task_id: Optional parent task ID
        task_id: Optional task ID being assigned
        
    Returns:
        Task ID if successful, None if failed
    """
    try:
        # Create task data with additional references
        task_data = {
            'resource_id': employee_id,
            'name': task_name,
            'start_datetime': task_start.strftime("%Y-%m-%d %H:%M:%S"),
            'end_datetime': task_end.strftime("%Y-%m-%d %H:%M:%S"),
            'role': 'Designer',  # Explicitly set role to Designer
        }
        
        # Add references to related Odoo tasks if available
        if task_id:
            task_data['x_studio_related_task_id'] = task_id
        
        if parent_task_id:
            task_data['x_studio_parent_task_id'] = parent_task_id
        
        task_id = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'planning.slot', 'create', [task_data]
        )
        
        logger.info(f"Created task in schedule (ID: {task_id})")
        return task_id
        
    except Exception as e:
        logger.error(f"Error creating task in schedule: {e}", exc_info=True)
        st.error(f"Error creating task in schedule: {e}")
        return None

def normalize_string(s: str) -> str:
    """
    Normalizes a string by removing special characters and extra spaces.
    
    Args:
        s: String to normalize
        
    Returns:
        Normalized string
    """
    return ''.join(e for e in s.lower().strip() if e.isalnum() or e.isspace()).replace("  ", " ")

def find_employee_id(employee_name: str, employees_in_planning: List[OdooRecord]) -> Optional[int]:
    """
    Finds an employee's ID by name.
    
    Args:
        employee_name: Name of the employee
        employees_in_planning: List of employees
        
    Returns:
        Employee ID if found, None otherwise
    """
    normalized_name = normalize_string(employee_name)
    for emp in employees_in_planning:
        if normalize_string(emp['name']) == normalized_name:
            return emp['id']
    
    # Try partial match if exact match fails
    for emp in employees_in_planning:
        if normalized_name in normalize_string(emp['name']):
            logger.info(f"Found partial match for employee: {employee_name} -> {emp['name']}")
            return emp['id']
            
    logger.warning(f"Could not find employee with name: {employee_name}")
    return None

def get_target_languages_odoo(models: xmlrpc.client.ServerProxy, uid: int) -> List[str]:
    """
    Retrieves a list of target languages from Odoo.
    
    Args:
        models: Odoo models proxy
        uid: User ID
        
    Returns:
        List of target languages
    """
    try:
        records = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'project.task', 'search_read',
            [[]],
            {'fields': ['x_studio_target_language']}
        )
        
        languages = set()
        for rec in records:
            lang = rec.get('x_studio_target_language')
            if lang:
                if isinstance(lang, list):
                    for l in lang:
                        languages.add(l)
                else:
                    languages.add(lang)
        
        logger.info(f"Retrieved {len(languages)} target languages")
        return sorted(list(languages))
        
    except Exception as e:
        logger.error(f"Error fetching target languages: {e}", exc_info=True)
        st.warning("Error fetching target languages, reauthenticating...")
        
        # Retry with new connection
        uid, models = authenticate_odoo()
        try:
            records = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                'project.task', 'search_read',
                [[]],
                {'fields': ['x_studio_target_language']}
            )
            
            languages = set()
            for rec in records:
                lang = rec.get('x_studio_target_language')
                if lang:
                    if isinstance(lang, list):
                        for l in lang:
                            languages.add(l)
                    else:
                        languages.add(lang)
            
            return sorted(list(languages))
        except Exception as e:
            logger.error(f"Retry failed to fetch target languages: {e}", exc_info=True)
            st.error(f"Failed to retrieve target languages: {e}")
            return []

def get_guidelines_odoo(models, uid):
    try:
        # First, get the fields available on the x_guidelines model
        field_info = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'x_guidelines', 'fields_get',
            [],
            {'attributes': ['string', 'type']}
        )
        
        # Find a suitable display field - it might be x_name or x_studio_name instead of 'name'
        display_field = 'name'  # Default attempt
        
        # Look for common name field alternatives
        for possible_field in ['x_name', 'x_studio_name', 'x_display_name']:
            if possible_field in field_info:
                display_field = possible_field
                break
                
        # If no standard name field found, use the first char/text field
        if display_field not in field_info:
            for field, info in field_info.items():
                if info.get('type') in ['char', 'text']:
                    display_field = field
                    break
        
        # Now fetch the guidelines with the correct field
        fields_to_fetch = ['id', display_field]
        guidelines_records = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'x_guidelines', 'search_read',
            [[]],
            {'fields': fields_to_fetch}
        )
        
        # Return as list of tuples (id, display_value)
        return [(rec['id'], rec.get(display_field, f"ID: {rec['id']}")) for rec in guidelines_records]
        
    except Exception as e:
        logger.error(f"Error fetching guidelines: {e}", exc_info=True)
        st.warning(f"Error fetching guidelines: {e}")
        return []

def get_client_success_executives_odoo(models: xmlrpc.client.ServerProxy, uid: int) -> List[OdooRecord]:
    """
    Retrieves a list of client success executives from Odoo.
    
    Args:
        models: Odoo models proxy
        uid: User ID
        
    Returns:
        List of client success executives
    """
    try:
        records = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'res.users', 'search_read',
            [[]],
            {'fields': ['id', 'name']}
        )
        
        logger.info(f"Retrieved {len(records)} client success executives")
        return records
        
    except Exception as e:
        logger.error(f"Error fetching client success executives: {e}", exc_info=True)
        st.warning("Error fetching client success executives, reauthenticating...")
        
        # Retry with new connection
        uid, models = authenticate_odoo()
        try:
            records = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                'res.users', 'search_read',
                [[]],
                {'fields': ['id', 'name']}
            )
            return records
        except Exception as e:
            logger.error(f"Retry failed to fetch client success executives: {e}", exc_info=True)
            st.error(f"Failed to retrieve client success executives: {e}")
            return []

def get_service_category_1_options(models: xmlrpc.client.ServerProxy, uid: int) -> List[Tuple[int, str]]:
    """
    Retrieves a list of service category 1 options from Odoo with their IDs.
    
    Args:
        models: Odoo models proxy
        uid: User ID
        
    Returns:
        List of tuples (id, name) for service category 1 options
    """
    try:
        # First, check if there's a dedicated model for service categories
        service_categories = []
        
        try:
            # Try several possible model names for service categories
            possible_models = ['x_service_category', 'x_studio_service_category', 'service.category']
            for model_name in possible_models:
                try:
                    # Try to query the model
                    category_records = models.execute_kw(
                        ODOO_DB, uid, ODOO_PASSWORD,
                        model_name, 'search_read',
                        [[]],
                        {'fields': ['id', 'name']}
                    )
                    if category_records:
                        service_categories = [(rec['id'], rec['name']) for rec in category_records]
                        logger.info(f"Retrieved {len(service_categories)} service categories from model {model_name}")
                        return service_categories
                except Exception:
                    # Continue to the next model name
                    continue
            
            logger.info("No dedicated service category model found, falling back to extraction from tasks")
        except Exception as e:
            logger.info(f"Error checking for dedicated service category models: {e}")
        
        # Fall back to extracting from existing tasks
        records = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'project.task', 'search_read',
            [[]],
            {'fields': ['id', 'x_studio_service_category_1']}
        )
        
        categories = set()
        for rec in records:
            cat = rec.get('x_studio_service_category_1')
            if not cat:
                continue
                
            # Handle different possible formats
            if isinstance(cat, list) and len(cat) == 2:
                # This is likely an [id, name] pair from Odoo
                categories.add((cat[0], cat[1]))
            elif isinstance(cat, int):
                # This is an ID reference
                try:
                    # Try to get the name from Odoo
                    category_name_records = models.execute_kw(
                        ODOO_DB, uid, ODOO_PASSWORD,
                        'x_service_category', 'read',  # Assuming this model exists
                        [[cat]],
                        {'fields': ['name']}
                    )
                    if category_name_records:
                        name = category_name_records[0].get('name', f"Category {cat}")
                        categories.add((cat, name))
                    else:
                        # If we can't get the name, use the ID as a string
                        categories.add((cat, f"Category {cat}"))
                except Exception:
                    # If the model doesn't exist or another error occurs
                    categories.add((cat, f"Category {cat}"))
            elif isinstance(cat, str):
                # For backward compatibility, create a dummy ID
                # This is not ideal but prevents immediate errors
                # The validate_task_data function should prevent this from being used
                logger.warning(f"Found string value for service category: {cat}")
                categories.add((-1, cat))  # Use -1 as a marker for invalid IDs
        
        logger.info(f"Retrieved {len(categories)} service category 1 options")
        return sorted(list(categories), key=lambda x: x[1])  # Sort by name
        
    except Exception as e:
        logger.error(f"Error fetching service category 1 options: {e}", exc_info=True)
        st.warning("Error fetching service categories, reauthenticating...")
        
        # Retry with new connection
        uid, models = authenticate_odoo()
        try:
            # Simple retry - just get base categories from tasks
            records = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                'project.task', 'search_read',
                [[]],
                {'fields': ['id', 'x_studio_service_category_1']}
            )
            
            categories = set()
            for rec in records:
                cat = rec.get('x_studio_service_category_1')
                if isinstance(cat, list) and len(cat) == 2:
                    categories.add((cat[0], cat[1]))
                elif isinstance(cat, int):
                    categories.add((cat, f"Category {cat}"))
                elif isinstance(cat, str) and cat:
                    categories.add((-1, cat))  # Use -1 as a marker for invalid IDs
            
            return sorted(list(categories), key=lambda x: x[1])
        except Exception as e:
            logger.error(f"Retry failed to fetch service categories: {e}", exc_info=True)
            st.error(f"Failed to retrieve service categories: {e}")
            return []

def get_service_category_2_options(models: xmlrpc.client.ServerProxy, uid: int) -> List[Tuple[int, str]]:
    """
    Retrieves a list of service category 2 options from Odoo with their IDs.
    
    Args:
        models: Odoo models proxy
        uid: User ID
        
    Returns:
        List of tuples (id, name) for service category 2 options
    """
    try:
        # First, check if there's a dedicated model for service categories
        service_categories = []
        
        try:
            # Try several possible model names for service categories
            possible_models = ['x_service_category_2', 'x_studio_service_category_2', 'service.category.2']
            for model_name in possible_models:
                try:
                    # Try to query the model
                    category_records = models.execute_kw(
                        ODOO_DB, uid, ODOO_PASSWORD,
                        model_name, 'search_read',
                        [[]],
                        {'fields': ['id', 'name']}
                    )
                    if category_records:
                        service_categories = [(rec['id'], rec['name']) for rec in category_records]
                        logger.info(f"Retrieved {len(service_categories)} service categories from model {model_name}")
                        return service_categories
                except Exception:
                    # Continue to the next model name
                    continue
            
            logger.info("No dedicated service category 2 model found, falling back to extraction from tasks")
        except Exception as e:
            logger.info(f"Error checking for dedicated service category 2 models: {e}")
        
        # Fall back to extracting from existing tasks
        records = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'project.task', 'search_read',
            [[]],
            {'fields': ['id', 'x_studio_service_category_2']}
        )
        
        categories = set()
        for rec in records:
            cat = rec.get('x_studio_service_category_2')
            if not cat:
                continue
                
            # Handle different possible formats
            if isinstance(cat, list) and len(cat) == 2:
                # This is likely an [id, name] pair from Odoo
                categories.add((cat[0], cat[1]))
            elif isinstance(cat, int):
                # This is an ID reference
                try:
                    # Try to get the name from Odoo
                    category_name_records = models.execute_kw(
                        ODOO_DB, uid, ODOO_PASSWORD,
                        'x_service_category_2', 'read',  # Assuming this model exists
                        [[cat]],
                        {'fields': ['name']}
                    )
                    if category_name_records:
                        name = category_name_records[0].get('name', f"Category {cat}")
                        categories.add((cat, name))
                    else:
                        # If we can't get the name, use the ID as a string
                        categories.add((cat, f"Category {cat}"))
                except Exception:
                    # If the model doesn't exist or another error occurs
                    categories.add((cat, f"Category {cat}"))
            elif isinstance(cat, str):
                # For backward compatibility, create a dummy ID
                # This is not ideal but prevents immediate errors
                logger.warning(f"Found string value for service category 2: {cat}")
                categories.add((-1, cat))  # Use -1 as a marker for invalid IDs
        
        logger.info(f"Retrieved {len(categories)} service category 2 options")
        return sorted(list(categories), key=lambda x: x[1])  # Sort by name
        
    except Exception as e:
        logger.error(f"Error fetching service category 2 options: {e}", exc_info=True)
        st.warning("Error fetching service category 2 options, reauthenticating...")
        
        # Retry with new connection
        uid, models = authenticate_odoo()
        try:
            # Simple retry - just get base categories from tasks
            records = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                'project.task', 'search_read',
                [[]],
                {'fields': ['id', 'x_studio_service_category_2']}
            )
            
            categories = set()
            for rec in records:
                cat = rec.get('x_studio_service_category_2')
                if isinstance(cat, list) and len(cat) == 2:
                    categories.add((cat[0], cat[1]))
                elif isinstance(cat, int):
                    categories.add((cat, f"Category {cat}"))
                elif isinstance(cat, str) and cat:
                    categories.add((-1, cat))  # Use -1 as a marker for invalid IDs
            
            return sorted(list(categories), key=lambda x: x[1])
        except Exception as e:
            logger.error(f"Retry failed to fetch service category 2 options: {e}", exc_info=True)
            st.error(f"Failed to retrieve service category 2 options: {e}")
            return []

def get_retainer_projects(models: xmlrpc.client.ServerProxy, uid: int, company_name: str = None) -> List[str]:
    """
    Retrieves a list of retainer projects from Odoo, optionally filtered by company.
    
    Args:
        models: Odoo models proxy
        uid: User ID
        company_name: Optional company name to filter by
        
    Returns:
        List of retainer project names
    """
    try:
        domain = []
        if company_name:
            domain = [('company_id.name', '=', company_name)]
            
        records = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'project.project', 'search_read',
            [domain],
            {'fields': ['name']}
        )
        
        project_names = [r['name'] for r in records if r.get('name')]
        logger.info(f"Retrieved {len(project_names)} retainer projects")
        return sorted(project_names)
        
    except Exception as e:
        logger.error(f"Error fetching retainer projects: {e}", exc_info=True)
        st.warning("Error fetching retainer projects, reauthenticating...")
        
        # Retry with new connection
        uid, models = authenticate_odoo()
        try:
            domain = []
            if company_name:
                domain = [('company_id.name', '=', company_name)]
                
            records = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                'project.project', 'search_read',
                [domain],
                {'fields': ['name']}
            )
            project_names = [r['name'] for r in records if r.get('name')]
            return sorted(project_names)
        except Exception as e:
            logger.error(f"Retry failed to fetch retainer projects: {e}", exc_info=True)
            st.error(f"Failed to retrieve retainer projects: {e}")
            return []

def get_retainer_customers(models: xmlrpc.client.ServerProxy, uid: int) -> List[str]:
    """
    Retrieves a list of retainer customers from Odoo.
    
    Args:
        models: Odoo models proxy
        uid: User ID
        
    Returns:
        List of retainer customer names
    """
    try:
        records = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'res.partner', 'search_read',
            [[['customer_rank', '>', 0]]],
            {'fields': ['name']}
        )
        
        customer_names = [r['name'] for r in records if r.get('name')]
        logger.info(f"Retrieved {len(customer_names)} retainer customers")
        return sorted(customer_names)
        
    except Exception as e:
        logger.error(f"Error fetching retainer customers: {e}", exc_info=True)
        st.warning("Error fetching retainer customers, reauthenticating...")
        
        # Retry with new connection
        uid, models = authenticate_odoo()
        try:
            records = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                'res.partner', 'search_read',
                [[['customer_rank', '>', 0]]],
                {'fields': ['name']}
            )
            customer_names = [r['name'] for r in records if r.get('name')]
            return sorted(customer_names)
        except Exception as e:
            logger.error(f"Retry failed to fetch retainer customers: {e}", exc_info=True)
            st.error(f"Failed to retrieve retainer customers: {e}")
            return []

def get_all_employees_in_planning(models: xmlrpc.client.ServerProxy, uid: int) -> List[OdooRecord]:
    """
    Retrieves a list of all employees in planning from Odoo.
    
    Args:
        models: Odoo models proxy
        uid: User ID
        
    Returns:
        List of employee records
    """
    try:
        records = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'resource.resource', 'search_read',
            [],
            {'fields': ['id', 'name']}
        )
        
        logger.info(f"Retrieved {len(records)} employees in planning")
        return records
        
    except Exception as e:
        logger.error(f"Error fetching employees in planning: {e}", exc_info=True)
        st.warning("Error fetching employees in planning, reauthenticating...")
        
        # Retry with new connection
        uid, models = authenticate_odoo()
        try:
            records = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                'resource.resource', 'search_read',
                [],
                {'fields': ['id', 'name']}
            )
            return records
        except Exception as e:
            logger.error(f"Retry failed to fetch employees in planning: {e}", exc_info=True)
            st.error(f"Failed to retrieve employees in planning: {e}")
            return []

def find_earliest_available_slot(schedule: List[Dict], task_duration: int, deadline: pd.Timestamp) -> Tuple[Optional[pd.Timestamp], Optional[pd.Timestamp]]:
    """
    Finds the earliest available time slot for a task.
    
    Args:
        schedule: List of scheduled tasks
        task_duration: Duration of the task in hours
        deadline: Deadline by which the task must be completed
        
    Returns:
        Tuple of (start_time, end_time) if found, (None, None) otherwise
    """
    try:
        now = pd.Timestamp.now().floor('min')
        task_duration_td = pd.Timedelta(hours=task_duration)
        
        if not isinstance(deadline, pd.Timestamp):
            deadline = pd.Timestamp(deadline)
        
        previous_end = now
        
        # Sort schedule by start time
        sorted_schedule = sorted(schedule, key=lambda x: pd.to_datetime(x['start_datetime']))

        for task in sorted_schedule:
            start = pd.to_datetime(task['start_datetime'])
            end = pd.to_datetime(task['end_datetime'])
            
            # Check if there's a gap before the next task
            if previous_end + task_duration_td <= start and previous_end + task_duration_td <= deadline:
                logger.info(f"Found available slot: {previous_end} - {previous_end + task_duration_td}")
                return previous_end, previous_end + task_duration_td
            
            previous_end = max(previous_end, end)

        # Check if there's room after all tasks
        if previous_end + task_duration_td <= deadline and previous_end >= now:
            logger.info(f"Found available slot after all tasks: {previous_end} - {previous_end + task_duration_td}")
            return previous_end, previous_end + task_duration_td

        logger.warning(f"No available slot found before deadline: {deadline}")
        return None, None
        
    except Exception as e:
        logger.error(f"Error finding available slot: {e}", exc_info=True)
        st.error(f"Error finding available time slot: {e}")
        return None, None

def get_project_id_by_name(models: xmlrpc.client.ServerProxy, uid: int, project_name: str) -> Optional[int]:
    """
    Gets the project ID by its name from Odoo.
    
    Args:
        models: Odoo models proxy
        uid: User ID
        project_name: Name of the project to find
        
    Returns:
        Project ID if found, None otherwise
    """
    try:
        projects = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'project.project', 'search_read',
            [[['name', '=', project_name]]],
            {'fields': ['id']}
        )
        
        if projects:
            project_id = projects[0]['id']
            logger.info(f"Found project ID {project_id} for project name: {project_name}")
            return project_id
        else:
            logger.warning(f"No project found with name: {project_name}")
            return None
            
    except Exception as e:
        logger.error(f"Error getting project ID by name: {e}", exc_info=True)
        st.warning(f"Error getting project ID, reauthenticating...")
        
        # Retry with new connection
        uid, models = authenticate_odoo()
        try:
            projects = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                'project.project', 'search_read',
                [[['name', '=', project_name]]],
                {'fields': ['id']}
            )
            
            if projects:
                project_id = projects[0]['id']
                return project_id
            else:
                return None
                
        except Exception as e:
            logger.error(f"Retry failed to get project ID: {e}", exc_info=True)
            st.error(f"Failed to get project ID: {e}")
            return None
        
# Add to helpers.py - New function to get companies
def get_companies(models: xmlrpc.client.ServerProxy, uid: int) -> List[str]:
    """
    Retrieves a list of all companies from Odoo.
    
    Args:
        models: Odoo models proxy
        uid: User ID
        
    Returns:
        List of company names
    """
    try:
        records = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'res.company', 'search_read',
            [[]],
            {'fields': ['id', 'name']}
        )
        
        company_names = [r['name'] for r in records if r.get('name')]
        logger.info(f"Retrieved {len(company_names)} companies")
        return sorted(company_names)
        
    except Exception as e:
        logger.error(f"Error fetching companies: {e}", exc_info=True)
        st.warning("Error fetching companies, reauthenticating...")
        
        # Retry with new connection
        uid, models = authenticate_odoo()
        try:
            records = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                'res.company', 'search_read',
                [[]],
                {'fields': ['id', 'name']}
            )
            company_names = [r['name'] for r in records if r.get('name')]
            return sorted(company_names)
        except Exception as e:
            logger.error(f"Retry failed to fetch companies: {e}", exc_info=True)
            st.error(f"Failed to retrieve companies: {e}")
            return []
        
def update_task_designer(models, uid, task_id, designer_name):
    """Simple version to identify what's working"""
    try:
        # Find employee ID
        employees = get_all_employees_in_planning(models, uid)
        employee_id = None
        for emp in employees:
            if designer_name.lower() in emp['name'].lower():
                employee_id = emp['id']
                break
        
        if not employee_id:
            logger.warning(f"Employee not found in planning: {designer_name}")
            return False
        
        # Update task with minimal information
        update_values = {
            'description': f"\n\nAssigned to designer: {designer_name} on {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        }
        
        # Try to find user ID
        user_ids = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'res.users', 'search_read',
            [[['name', 'ilike', designer_name]]],
            {'fields': ['id', 'name']}
        )
        
        if user_ids:
            update_values['user_id'] = user_ids[0]['id']
            
        # Log what we're updating
        logger.info(f"Updating task {task_id} with values: {update_values}")
        
        # Update the task
        result = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'project.task', 'write',
            [[task_id], update_values]
        )
        
        return bool(result)
    except Exception as e:
        logger.error(f"Error updating task with designer: {e}", exc_info=True)
        return False

def test_designer_update(models, uid, task_id):
    """
    Test function that makes a minimal change to a task
    to debug designer assignment issues
    """
    try:
        # Get the current logged-in user (which should have permissions)
        # and assign them to the task as a test
        logger.info(f"Testing task update with current user (uid={uid})")
        
        # Try to update the task with the current user
        result = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'project.task', 'write',
            [[task_id], {'user_id': uid}]
        )
        
        if result:
            logger.info("Test update successful!")
            return True
        else:
            logger.error("Test update failed but no error was thrown")
            return False
    except Exception as e:
        logger.error(f"Test update failed with error: {e}")
        return False
    
# Add this debugging function to helpers.py
def get_available_fields(models, uid, model_name='planning.slot'):
    """Get all available fields for a model"""
    try:
        fields = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            model_name, 'fields_get',
            [], {'attributes': ['string', 'type', 'required']}
        )
        return fields
    except Exception as e:
        logger.error(f"Error getting fields for {model_name}: {e}")
        return {}