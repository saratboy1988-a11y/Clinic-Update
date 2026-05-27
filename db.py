import sqlite3
import hashlib
import hmac
import os
import sys
import logging
import secrets
import string
import ctypes
from datetime import datetime

# កំណត់ទីតាំង Database ឱ្យនៅជាមួយ File កម្មវិធីជានិច្ច
# អាចប្តូរទីតាំងតាមរយៈ Environment Variable: CLINIC_DB_PATH

# Priority for database location:
# 1. Environment variable CLINIC_DB_PATH (highest priority)
# 2. AppData folder (for installed apps in Program Files)
# 3. Same directory as executable/script
# 4. Documents folder (fallback)

import ctypes.wintypes

# Check if user has specified a custom database path
custom_db_path = os.getenv("CLINIC_DB_PATH")

if custom_db_path:
    # Use the custom path from environment variable
    base_dir = os.path.dirname(custom_db_path)
    DB_NAME = custom_db_path
    logger_info = f"Using custom database path from environment variable"
else:
    # For frozen app, use the directory where the executable is located
    # For development, use the script directory
    if getattr(sys, 'frozen', False):
        # When running as compiled executable
        base_dir = os.path.dirname(sys.executable)
        
        # Check if running from Program Files (needs admin rights)
        # If so, use AppData instead
        program_files = os.getenv("ProgramFiles")
        if program_files and base_dir.startswith(program_files):
            # Use AppData for database when installed in Program Files
            CSIDL_APPDATA = 26  # Roaming AppData
            SHGFP_TYPE_CURRENT = 0
            buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
            ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_APPDATA, None, SHGFP_TYPE_CURRENT, buf)
            base_dir = os.path.join(buf.value, "ClinicManager")
            if not os.path.exists(base_dir):
                os.makedirs(base_dir)
            logger_info = f"Using AppData folder (Program Files installation)"
        else:
            logger_info = f"Using executable directory"
    else:
        # When running as script
        base_dir = os.path.dirname(os.path.abspath(__file__))
        logger_info = f"Using script directory"

    DB_NAME = os.path.join(base_dir, "clinic.db")

# Configure logging to use the same base directory
log_file = os.path.join(base_dir, "clinic.log")

# Ensure the directory exists and is writable
try:
    if not os.path.exists(base_dir):
        os.makedirs(base_dir)
        logger_info = f"Created directory: {base_dir}"
    
    # Test if directory is writable (skip if using custom path - user's responsibility)
    if not custom_db_path:
        test_file = os.path.join(base_dir, ".write_test")
        try:
            with open(test_file, 'w') as f:
                f.write("test")
            os.remove(test_file)
        except (IOError, OSError):
            # If not writable, use user's documents folder
            CSIDL_PERSONAL = 5  # My Documents
            SHGFP_TYPE_CURRENT = 0
            buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
            ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
            base_dir = os.path.join(buf.value, "ClinicManager")
            if not os.path.exists(base_dir):
                os.makedirs(base_dir)
            DB_NAME = os.path.join(base_dir, "clinic.db")
            log_file = os.path.join(base_dir, "clinic.log")
            logger_info = f"Fallback to Documents folder: {base_dir}"
except Exception as e:
    # Fallback to current directory
    base_dir = os.getcwd()
    DB_NAME = os.path.join(base_dir, "clinic.db")
    log_file = os.path.join(base_dir, "clinic.log")
    logger_info = f"Fallback to current directory: {base_dir}"

# Setup logging with both file and console handlers
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# File handler
try:
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    ))
    logger.addHandler(file_handler)
except Exception:
    # If can't write to file, skip file logging
    pass

# Console handler
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(logging.Formatter(
    '%(asctime)s - %(levelname)s - %(message)s'
))
logger.addHandler(console_handler)

class PatientCol:
    ID = 0
    DATE = 1
    SERIAL = 2
    CARD_ID = 3
    NAME = 4
    GUARDIAN = 5
    AGE = 6
    SEX = 7
    AREA = 8
    PREGNANT = 9
    ADDRESS = 10
    PHONE = 11
    REF_FROM = 12
    DISEASE = 13
    SYMPTOMS = 14
    PARACLINICAL = 15
    DIAGNOSIS = 16
    TREATMENT = 17
    IMCI = 18
    NUTRITION = 19
    REF_TO = 20
    SERVICE = 21
    REMARK = 22
    TYPE = 23
    BRANCH = 24


PATIENT_COLUMNS = {
    "date", "serial_no", "card_id", "name", "guardian", "age", "sex", "area",
    "pregnant", "address", "phone", "ref_from", "disease_case", "symptoms",
    "paraclinical", "diagnosis", "treatment", "imci", "nutrition", "ref_to",
    "service", "remark", "patient_type", "branch_code"
}

# Helper function to handle database connection and write operations (Insert, Update, Delete)
def execute_write(query, params=(), return_lastrowid=False):
    """
    Execute a write query on the database

    Args:
        query: SQL query to execute
        params: Query parameters
        return_lastrowid: If True, returns the last inserted rowid

    Returns:
        int or None: The last inserted rowid if return_lastrowid is True, else None
    """
    try:
        with sqlite3.connect(DB_NAME) as conn:
            cur = conn.cursor()
            cur.execute(query, params)
            conn.commit()
            if return_lastrowid:
                return cur.lastrowid
    except sqlite3.Error as e:
        logger.error(f"Database write error: {e} - Query: {query}")
        raise
    return None


def _patient_insert_columns():
    return (
        "date", "serial_no", "card_id", "name", "guardian", "age", "sex", "area", "pregnant",
        "address", "phone", "ref_from", "disease_case", "symptoms", "paraclinical", "diagnosis",
        "treatment", "imci", "nutrition", "ref_to", "service", "remark", "patient_type", "branch_code"
    )


def normalize_branch_code(branch_code):
    branch = str(branch_code or "").strip().upper()
    return branch or "MAIN"


def is_admin_user_row(user):
    if not user:
        return False
    user_dict = dict(user)
    return str(user_dict.get("role", "")).strip().lower() == "admin" or str(user_dict.get("username", "")).strip().lower() == "admin"


def _branch_filter_clause(branch_code, column_name="branch_code"):
    branch = normalize_branch_code(branch_code)
    if branch == "ALL":
        return "1=1", []
    return f"({column_name} = ? OR {column_name} IS NULL OR {column_name} = '')", [branch]


def _insert_patient_with_cursor(cur, patient_data):
    patient_data = list(patient_data)
    columns = _patient_insert_columns()
    type_idx = columns.index("patient_type")
    branch_idx = columns.index("branch_code")
    if len(patient_data) > type_idx:
        patient_data[type_idx] = _infer_patient_type(
            patient_data[PatientCol.SERIAL - 1],
            patient_data[type_idx],
            patient_data[PatientCol.AGE - 1],
        )
    if len(patient_data) < len(columns):
        patient_data.append("MAIN")
    patient_data[branch_idx] = normalize_branch_code(patient_data[branch_idx])
    patient_data = tuple(patient_data)
    placeholders = ",".join(["?"] * len(columns))
    query = f"INSERT INTO patient ({','.join(columns)}) VALUES ({placeholders})"
    cur.execute(query, patient_data)


def _normalize_merge_text(value):
    return str(value or "").strip()


def _build_merge_key(date, name, card_id, guardian, patient_type="", branch_code=""):
    date_key = _normalize_merge_text(date)
    name_key = _normalize_merge_text(name).lower()
    card_key = _normalize_merge_text(card_id).lower()
    guardian_key = _normalize_merge_text(guardian).lower()
    type_key = _normalize_patient_type(patient_type).lower()
    branch_key = normalize_branch_code(branch_code)
    return f"{date_key}|{name_key}|{card_key}|{guardian_key}|{type_key}|{branch_key}"


def _normalize_patient_type(patient_type, serial_no=""):
    normalized_type = _normalize_merge_text(patient_type)
    normalized_lower = normalized_type.lower()

    if normalized_lower in ("adult", "a", "a-", "មនុស្សចាស់", "មនុស្សធំ"):
        return "Adult"
    if normalized_lower in ("child", "c", "c-", "កុមារ"):
        return "Child"
    if "adult" in normalized_lower:
        return "Adult"
    if "child" in normalized_lower:
        return "Child"

    serial_text = _normalize_merge_text(serial_no)
    if serial_text.startswith(("A-", "a-", "A", "a")):
        return "Adult"
    return "Child"


def _infer_patient_type_from_age(age_text):
    age_category = _normalize_merge_text(age_text).split("::", 1)[0]
    child_age_groups = ("០-២៩ថ្ងៃ", "០-២៨ថ្ងៃ", "២៩ថ្ងៃ-១១ខែ", "១-៤ឆ្នាំ", "0-28", "0-29", "29-11", "1-4")
    adult_age_groups = ("៥-១៤ឆ្នាំ", "១៥-២៤ឆ្នាំ", "២៥-៤៩ឆ្នាំ", "៥០-៦៤ឆ្នាំ", ">=៦៥ឆ្នាំ", ">=៦៤ឆ្នាំ",
                        "5-14", "15-24", "25-49", "50-64", ">=65", ">=64")

    if any(group in age_category for group in child_age_groups):
        return "Child"
    if any(group in age_category for group in adult_age_groups):
        return "Adult"
    return None


def _infer_patient_type(serial_no, patient_type, age_text=""):
    age_type = _infer_patient_type_from_age(age_text)
    if age_type:
        return age_type
    return _normalize_patient_type(patient_type, serial_no)


def _date_sort_key(date_text):
    try:
        return datetime.strptime(_normalize_merge_text(date_text), "%d/%m/%Y")
    except (TypeError, ValueError):
        return datetime.max


def _month_key(date_text):
    try:
        return datetime.strptime(_normalize_merge_text(date_text), "%d/%m/%Y").strftime("%Y%m")
    except (TypeError, ValueError):
        return ""


def _parse_serial_number(serial_no, prefix):
    serial_text = _normalize_merge_text(serial_no)
    try:
        if serial_text.lower().startswith(prefix.lower()):
            return int(serial_text[len(prefix):])
        return int(serial_text)
    except (TypeError, ValueError):
        return 0


def _get_max_serials_by_type_month_with_cursor(cur, branch_code=None):
    branch_clause, branch_params = _branch_filter_clause(branch_code) if branch_code else ("1=1", [])
    cur.execute(f"SELECT date, serial_no, patient_type FROM patient WHERE {branch_clause}", tuple(branch_params))
    max_by_type_month = {}
    for date_text, serial_no, patient_type in cur.fetchall():
        normalized_type = _normalize_patient_type(patient_type, serial_no)
        prefix = "A-" if normalized_type == "Adult" else "C-"
        key = (normalized_type, _month_key(date_text))
        max_by_type_month[key] = max(max_by_type_month.get(key, 0), _parse_serial_number(serial_no, prefix))
    return max_by_type_month


def _assign_continuing_serials_for_merge(cur, patients, branch_code=None):
    max_by_type_month = _get_max_serials_by_type_month_with_cursor(cur, branch_code)
    assigned = []

    for patient in sorted(patients, key=lambda p: (_date_sort_key(p[0]), p[3].lower(), p[2].lower())):
        patient_list = list(patient)
        patient_type = _infer_patient_type(patient_list[1], patient_list[22], patient_list[5])
        key = (patient_type, _month_key(patient_list[0]))
        max_by_type_month[key] = max_by_type_month.get(key, 0) + 1
        prefix = "A-" if patient_type == "Adult" else "C-"
        patient_list[1] = f"{prefix}{max_by_type_month[key]}"
        patient_list[22] = patient_type
        assigned.append(tuple(patient_list))

    return assigned


def _build_patient_tuple_from_source_row(row, source_columns):
    column_index = {column: idx for idx, column in enumerate(source_columns)}

    def get_value(column_name, default=""):
        idx = column_index.get(column_name)
        if idx is None or idx >= len(row):
            return default
        value = row[idx]
        return default if value is None else value

    serial_no = get_value("serial_no", "")
    patient_type = _infer_patient_type(serial_no, get_value("patient_type", ""), get_value("age", ""))

    return (
        get_value("date", ""),
        serial_no,
        get_value("card_id", ""),
        get_value("name", ""),
        get_value("guardian", ""),
        get_value("age", ""),
        get_value("sex", ""),
        get_value("area", ""),
        get_value("pregnant", ""),
        get_value("address", ""),
        get_value("phone", ""),
        get_value("ref_from", ""),
        get_value("disease_case", ""),
        get_value("symptoms", ""),
        get_value("paraclinical", ""),
        get_value("diagnosis", ""),
        get_value("treatment", ""),
        get_value("imci", ""),
        get_value("nutrition", ""),
        get_value("ref_to", ""),
        get_value("service", ""),
        get_value("remark", ""),
        patient_type,
        normalize_branch_code(get_value("branch_code", "")),
    )


def _renumber_serials_with_cursor(cur):
    cur.execute("SELECT id, patient_type, serial_no FROM patient")
    type_updates = [
        (_normalize_patient_type(patient_type, serial_no), patient_id)
        for patient_id, patient_type, serial_no in cur.fetchall()
    ]
    if type_updates:
        cur.executemany("UPDATE patient SET patient_type=? WHERE id=?", type_updates)

    cur.execute(
        """
        SELECT id, date
        FROM patient
        WHERE patient_type = 'Child' OR patient_type IS NULL OR patient_type = ''
        ORDER BY SUBSTR(date, 7, 4), SUBSTR(date, 4, 2), SUBSTR(date, 1, 2), id ASC
        """
    )
    child_counters = {}
    for patient_id, date_text in cur.fetchall():
        key = _month_key(date_text)
        child_counters[key] = child_counters.get(key, 0) + 1
        cur.execute("UPDATE patient SET serial_no=? WHERE id=?", (f"C-{child_counters[key]}", patient_id))

    cur.execute(
        """
        SELECT id, date
        FROM patient
        WHERE patient_type = 'Adult'
        ORDER BY SUBSTR(date, 7, 4), SUBSTR(date, 4, 2), SUBSTR(date, 1, 2), id ASC
        """
    )
    adult_counters = {}
    for patient_id, date_text in cur.fetchall():
        key = _month_key(date_text)
        adult_counters[key] = adult_counters.get(key, 0) + 1
        cur.execute("UPDATE patient SET serial_no=? WHERE id=?", (f"A-{adult_counters[key]}", patient_id))


def _renumber_categories_with_cursor(cur):
    patient_queries = (
        ("Child", "SELECT id, area, disease_case, imci, patient_type FROM patient WHERE patient_type = 'Child' OR patient_type IS NULL OR patient_type = '' ORDER BY id ASC"),
        ("Adult", "SELECT id, area, disease_case, imci, patient_type FROM patient WHERE patient_type = 'Adult' ORDER BY id ASC"),
    )

    for patient_type, query in patient_queries:
        cur.execute(query)
        rows = cur.fetchall()
        if not rows:
            continue

        area_counters = {}
        disease_counters = {}
        imci_counters = {}
        updates = []

        for patient_id, area_raw, disease_raw, imci_raw, row_patient_type in rows:
            area_cat = str(area_raw or "")
            area_cat = area_cat.split("::")[0] if "::" in area_cat else area_cat
            new_area = ""
            if area_cat:
                area_counters[area_cat] = area_counters.get(area_cat, 0) + 1
                new_area = f"{area_cat}::{area_counters[area_cat]}"

            disease_cat = str(disease_raw or "")
            disease_cat = disease_cat.split("::")[0] if "::" in disease_cat else disease_cat
            new_disease = ""
            if disease_cat:
                disease_counters[disease_cat] = disease_counters.get(disease_cat, 0) + 1
                new_disease = f"{disease_cat}::{disease_counters[disease_cat]}"

            imci_cat = str(imci_raw or "")
            imci_cat = imci_cat.split("::")[0] if "::" in imci_cat else imci_cat
            imci_cat = imci_cat.strip()  # Remove whitespace
            
            # Auto-assign IMCI "Yes" for all children if empty
            if patient_type == "Child" and (not imci_cat or imci_cat.lower() in ['', 'none', 'n/a']):
                imci_cat = "Yes"
            
            if patient_type == "Child" and imci_cat:
                imci_counters[imci_cat] = imci_counters.get(imci_cat, 0) + 1
                new_imci = f"{imci_cat}::{imci_counters[imci_cat]}"
            else:
                new_imci = imci_cat if patient_type == "Adult" else ""

            if patient_type == "Adult":
                new_imci = str(imci_raw or "")

            updates.append((new_area, new_disease, new_imci, patient_id))

        cur.executemany("UPDATE patient SET area=?, disease_case=?, imci=? WHERE id=?", updates)


def renumber_all_patient_numbering():
    """Normalize patient type without changing existing patient serial/code values."""
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, patient_type, serial_no FROM patient")
        type_updates = [
            (_normalize_patient_type(patient_type, serial_no), patient_id)
            for patient_id, patient_type, serial_no in cur.fetchall()
        ]
        if type_updates:
            cur.executemany("UPDATE patient SET patient_type=? WHERE id=?", type_updates)
        conn.commit()


def merge_patients(source_patients, source_columns, branch_code=None):
    """
    Merge patient rows from an external database using one duplicate rule and one transaction.

    Returns:
        tuple[int, int]: (merged_count, skipped_count)
    """
    if not source_patients:
        return 0, 0

    merged_count = 0
    skipped_count = 0

    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute("SELECT date, name, card_id, COALESCE(guardian, ''), patient_type, branch_code FROM patient")
        existing_keys = {
            _build_merge_key(row[0], row[1], row[2], row[3], row[4], row[5])
            for row in cur.fetchall()
        }

        patients_to_insert = []
        for row in source_patients:
            patient_list = list(_build_patient_tuple_from_source_row(row, source_columns))
            if branch_code:
                patient_list[-1] = normalize_branch_code(branch_code)
            patient = tuple(patient_list)
            merge_key = _build_merge_key(patient[0], patient[3], patient[2], patient[4], patient[22], patient[23])
            if merge_key in existing_keys:
                skipped_count += 1
                continue

            patients_to_insert.append(patient)
            existing_keys.add(merge_key)
            merged_count += 1

        for patient in _assign_continuing_serials_for_merge(cur, patients_to_insert, branch_code):
            _insert_patient_with_cursor(cur, patient)

        conn.commit()

    return merged_count, skipped_count


def merge_database_file(source_db_path, branch_code=None):
    """
    Merge patient data from a source SQLite database into the current database.

    Returns:
        tuple[int, int]: (merged_count, skipped_count)
    """
    try:
        with sqlite3.connect(source_db_path) as source_conn:
            source_cur = source_conn.cursor()
            source_cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='patient'")
            if not source_cur.fetchone():
                return 0, 0

            source_cur.execute("PRAGMA table_info(patient)")
            columns = [row[1] for row in source_cur.fetchall()]
            source_cur.execute("SELECT * FROM patient")
            source_patients = source_cur.fetchall()

        return merge_patients(source_patients, columns, branch_code)
    except Exception as e:
        logger.error(f"Merge error: {e}")
        raise

# Helper function to handle database connection and read operations (Select)
def execute_read(query, params=(), one=False):
    try:
        with sqlite3.connect(DB_NAME) as conn:
            conn.row_factory = sqlite3.Row  # អនុញ្ញាតឱ្យ access តាមឈ្មោះ column
            cur = conn.cursor()
            cur.execute(query, params)
            rv = cur.fetchone() if one else cur.fetchall()
            return rv
    except sqlite3.Error as e:
        logger.error(f"Database read error: {e} - Query: {query}")
        return None if one else []

# Helper function to hash passwords
PASSWORD_HASH_ITERATIONS = 260000
PASSWORD_HASH_PREFIX = "pbkdf2_sha256"


def _hash_password(password):
    salt = secrets.token_hex(16)
    derived_key = hashlib.pbkdf2_hmac(
        'sha256',
        password.encode('utf-8'),
        salt.encode('utf-8'),
        PASSWORD_HASH_ITERATIONS
    )
    return f"{PASSWORD_HASH_PREFIX}${PASSWORD_HASH_ITERATIONS}${salt}${derived_key.hex()}"


def _verify_password(password, stored_password):
    stored_password = str(stored_password or "")

    if stored_password.startswith(f"{PASSWORD_HASH_PREFIX}$"):
        try:
            _, iterations_text, salt, expected_hash = stored_password.split("$", 3)
            iterations = int(iterations_text)
        except ValueError:
            return False, False

        derived_key = hashlib.pbkdf2_hmac(
            'sha256',
            password.encode('utf-8'),
            salt.encode('utf-8'),
            iterations
        ).hex()
        return hmac.compare_digest(derived_key, expected_hash), False

    legacy_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
    return hmac.compare_digest(legacy_hash, stored_password), True


def _build_patient_type_filter(patient_type):
    """
    Returns SQL and params for patient type filtering while keeping
    backward compatibility with legacy rows that have NULL/empty type.
    """
    if patient_type == 'Adult':
        return "patient_type = ?", ['Adult']
    if patient_type == 'Child':
        return "(patient_type = ? OR patient_type IS NULL OR patient_type = '')", ['Child']
    return "1=1", []


def bulk_replace_patients(patients, branch_code=None):
    """Replace all patient rows atomically."""
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        if branch_code:
            branch_clause, branch_params = _branch_filter_clause(branch_code)
            cur.execute(f"DELETE FROM patient WHERE {branch_clause}", tuple(branch_params))
        else:
            cur.execute("DELETE FROM patient")
        for patient in patients:
            patient_list = list(patient)
            if branch_code:
                if len(patient_list) < len(_patient_insert_columns()):
                    patient_list.append(branch_code)
                patient_list[-1] = normalize_branch_code(branch_code)
            _insert_patient_with_cursor(cur, tuple(patient_list))
        conn.commit()


def bulk_insert_patients(patients, branch_code=None):
    """Insert many patient rows atomically."""
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        for patient in patients:
            patient_list = list(patient)
            if branch_code:
                if len(patient_list) < len(_patient_insert_columns()):
                    patient_list.append(branch_code)
                patient_list[-1] = normalize_branch_code(branch_code)
            _insert_patient_with_cursor(cur, tuple(patient_list))
        conn.commit()


def _generate_initial_admin_password(length=12):
    alphabet = string.ascii_letters + string.digits
    return ''.join(secrets.choice(alphabet) for _ in range(length))

def connect():
    """
    Initialize database connection and create tables if they don't exist.
    Includes error handling for disk I/O issues.
    """
    global base_dir, DB_NAME
    
    try:
        # Ensure directory exists and is writable
        db_dir = os.path.dirname(DB_NAME)
        
        # Try to create directory if it doesn't exist
        if db_dir and not os.path.exists(db_dir):
            try:
                os.makedirs(db_dir)
                logger.info(f"Created database directory: {db_dir}")
            except Exception as e:
                logger.error(f"Failed to create directory {db_dir}: {e}")
                # Fallback to AppData
                CSIDL_APPDATA = 26
                SHGFP_TYPE_CURRENT = 0
                buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
                ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_APPDATA, None, SHGFP_TYPE_CURRENT, buf)
                base_dir = os.path.join(buf.value, "ClinicManager")
                if not os.path.exists(base_dir):
                    os.makedirs(base_dir)
                DB_NAME = os.path.join(base_dir, "clinic.db")
                log_file = os.path.join(base_dir, "clinic.log")
                logger.info(f"Fallback to AppData: {base_dir}")
        
        # Test if directory is writable
        test_file = os.path.join(db_dir, ".test_write")
        try:
            with open(test_file, 'w') as f:
                f.write("test")
            os.remove(test_file)
            logger.info(f"Directory is writable: {db_dir}")
        except Exception as e:
            logger.error(f"Directory not writable: {e}")
            # Fallback to AppData
            CSIDL_APPDATA = 26
            SHGFP_TYPE_CURRENT = 0
            buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
            ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_APPDATA, None, SHGFP_TYPE_CURRENT, buf)
            base_dir = os.path.join(buf.value, "ClinicManager")
            if not os.path.exists(base_dir):
                os.makedirs(base_dir)
            DB_NAME = os.path.join(base_dir, "clinic.db")
            log_file = os.path.join(base_dir, "clinic.log")
            logger.info(f"Fallback to AppData: {base_dir}")
        
        logger.info(f"Initializing database at: {DB_NAME}")
        logger.info(f"Database directory: {base_dir}")
        logger.info(f"Configuration: {logger_info}")

        # Connect to database with retry logic
        conn = None
        max_retries = 3
        for attempt in range(max_retries):
            try:
                conn = sqlite3.connect(DB_NAME)
                logger.info(f"Database connection successful on attempt {attempt + 1}")
                break
            except sqlite3.Error as e:
                logger.warning(f"Connection attempt {attempt + 1} failed: {e}")
                if attempt == max_retries - 1:
                    logger.error(f"All {max_retries} connection attempts failed")
                    raise
                import time
                time.sleep(0.5)  # Wait before retry
        
        if conn:
            cur = conn.cursor()

            # --- ១. បង្កើតតារាងជាមុនសិន (Create Tables First) ---
            cur.execute("""
                CREATE TABLE IF NOT EXISTS patient(
                    id INTEGER PRIMARY KEY,
                    date TEXT,
                    serial_no TEXT,
                    card_id TEXT,
                    name TEXT,
                    guardian TEXT,
                    age TEXT,
                    sex TEXT,
                    area TEXT,
                    pregnant TEXT,
                    address TEXT,
                    phone TEXT,
                    ref_from TEXT,
                    disease_case TEXT,
                    symptoms TEXT,
                    paraclinical TEXT,
                    diagnosis TEXT,
                    treatment TEXT,
                    imci TEXT,
                    nutrition TEXT,
                    ref_to TEXT,
                    service TEXT,
                    remark TEXT,
                    patient_type TEXT,
                    branch_code TEXT
                )
            """)
            # បន្ថែម Index ដើម្បីឱ្យការ Search លឿនជាងមុន
            cur.execute("CREATE INDEX IF NOT EXISTS idx_patient_name ON patient(name)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_patient_serial ON patient(serial_no)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_patient_card ON patient(card_id)")

            cur.execute("CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, username TEXT, password TEXT, role TEXT, branch_code TEXT)")
            cur.execute("CREATE TABLE IF NOT EXISTS login_history(id INTEGER PRIMARY KEY, username TEXT, login_time TEXT)")
            cur.execute("""
                CREATE TABLE IF NOT EXISTS license(
                    id INTEGER PRIMARY KEY, install_date TEXT, license_key TEXT, email TEXT, machine_id TEXT
                )
            """)

            # --- Safe Schema Migration using ALTER TABLE ---
            # Define the target schema to ensure all columns exist
            target_schema = {
                "id": "INTEGER PRIMARY KEY", "date": "TEXT", "serial_no": "TEXT", "card_id": "TEXT",
                "name": "TEXT", "guardian": "TEXT", "age": "TEXT", "sex": "TEXT", "area": "TEXT",
                "pregnant": "TEXT", "address": "TEXT", "phone": "TEXT", "ref_from": "TEXT",
                "disease_case": "TEXT", "symptoms": "TEXT", "paraclinical": "TEXT", "diagnosis": "TEXT",
                "treatment": "TEXT", "imci": "TEXT", "nutrition": "TEXT", "ref_to": "TEXT",
                "service": "TEXT", "remark": "TEXT", "patient_type": "TEXT", "branch_code": "TEXT"
            }

            # Get existing columns from the patient table
            cur.execute("PRAGMA table_info(patient)")
            existing_columns = {row[1] for row in cur.fetchall()}

            # Add any missing columns without dropping the table
            for col_name, col_type in target_schema.items():
                if col_name not in existing_columns:
                    logger.info(f"Schema migration: Adding missing column '{col_name}' to 'patient' table.")
                    cur.execute(f"ALTER TABLE patient ADD COLUMN {col_name} {col_type}")

            cur.execute("UPDATE patient SET branch_code='MAIN' WHERE branch_code IS NULL OR branch_code=''")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_patient_branch ON patient(branch_code)")

            cur.execute("PRAGMA table_info(users)")
            existing_user_columns = {row[1] for row in cur.fetchall()}
            if "role" not in existing_user_columns:
                cur.execute("ALTER TABLE users ADD COLUMN role TEXT")
            if "branch_code" not in existing_user_columns:
                cur.execute("ALTER TABLE users ADD COLUMN branch_code TEXT")
            cur.execute("UPDATE users SET role='user' WHERE role IS NULL OR role=''")
            cur.execute("UPDATE users SET branch_code='MAIN' WHERE branch_code IS NULL OR branch_code=''")
            cur.execute("UPDATE users SET role='admin', branch_code='ALL' WHERE username='admin'")

            # Enforce unique usernames when the existing data allows it.
            cur.execute("""
                SELECT username, COUNT(*)
                FROM users
                GROUP BY username
                HAVING COUNT(*) > 1
            """)
            duplicate_users = cur.fetchall()
            if duplicate_users:
                logger.warning("Skipping unique index on users.username because duplicates already exist")
            else:
                cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_unique ON users(username)")

            # Create default admin user if not exists
            cur.execute("SELECT * FROM users WHERE username = 'admin'")
            if not cur.fetchone():
                initial_admin_password = os.getenv("CLINIC_INITIAL_ADMIN_PASSWORD") or _generate_initial_admin_password()
                hashed_admin_pass = _hash_password(initial_admin_password)
                cur.execute(
                    "INSERT INTO users (username, password, role, branch_code) VALUES ('admin', ?, 'admin', 'ALL')",
                    (hashed_admin_pass,)
                )
                logger.warning("Created default admin user with a generated initial password")
                try:
                    ctypes.windll.user32.MessageBoxW(
                        None,
                        "A new admin account was created.\n\n"
                        "Username: admin\n"
                        f"Temporary Password: {initial_admin_password}\n\n"
                        "Please sign in and change this password immediately.",
                        "Initial Admin Account",
                        64
                    )
                except Exception:
                    logger.warning(
                        "Created default admin user, but the temporary password dialog could not be displayed."
                    )

            conn.commit()
            logger.info("Database initialization completed successfully")
            conn.close()

    except sqlite3.Error as e:
        error_msg = f"Database initialization error: {e}"
        logger.error(error_msg)
        logger.error(f"Database path: {DB_NAME}")
        logger.error(f"Current directory: {os.getcwd()}")
        logger.error(f"Base directory: {base_dir}")
        # Show user-friendly error
        ctypes.windll.user32.MessageBoxW(
            None,
            f"មិនអាចបើក Database បានទេ!\n\n" +
            f"កំហុស: {str(e)}\n\n" +
            f"ទីតាំង: {DB_NAME}\n\n" +
            f"សូមពិនិត្យមើលថា៖\n" +
            f"1. Antivirus មិនបានចាប់កម្មវិធី\n" +
            f"2. មានសិទ្ធិសរសេរក្នុង folder នេះ\n" +
            f"3. Disk មិនពេញ\n",
            "Database Error",
            16  # MB_ICONERROR
        )
        raise
    except Exception as e:
        logger.error(f"Unexpected error during database initialization: {e}")
        raise

def get_license_info():
    return execute_read("SELECT * FROM license LIMIT 1", one=True)

def get_patient_by_id(patient_id):
    """
    Get a single patient record by ID

    Args:
        patient_id: Patient ID to retrieve

    Returns:
        Tuple containing patient data or None if not found
    """
    return execute_read("SELECT * FROM patient WHERE id=?", (patient_id,), one=True)


def get_user_context(username):
    user = execute_read("SELECT * FROM users WHERE username=?", (username,), one=True)
    if not user:
        return {"username": username, "role": "user", "branch_code": "MAIN", "is_admin": False}
    user_dict = dict(user)
    role = str(user_dict.get("role") or "user").strip().lower()
    branch_code = normalize_branch_code(user_dict.get("branch_code") or ("ALL" if str(username).lower() == "admin" else "MAIN"))
    is_admin = role == "admin" or str(username).strip().lower() == "admin"
    if is_admin:
        branch_code = "ALL"
    return {
        "username": user_dict.get("username", username),
        "role": "admin" if is_admin else "user",
        "branch_code": branch_code,
        "is_admin": is_admin,
    }

def save_license_info(install_date, license_key, email, machine_id):
    execute_write("DELETE FROM license")
    execute_write("INSERT INTO license VALUES (NULL, ?, ?, ?, ?)", (install_date, license_key, email, machine_id))

def insert(date, serial_no, card_id, name, guardian, age, sex, area, pregnant, address, phone,
           ref_from, disease_case, symptoms, paraclinical, diagnosis, treatment,
           imci, nutrition, ref_to, service, remark, patient_type, id=None, branch_code="MAIN"):
    """
    Insert a new patient record (or update if id is provided)

    Args:
        id: Optional patient ID (if provided, uses INSERT OR REPLACE)
        All other args: Patient data fields

    Returns:
        int: The inserted/updated patient ID
    """
    patient_type = _infer_patient_type(serial_no, patient_type, age)
    branch_code = normalize_branch_code(branch_code)
    if id is not None:
        # Insert with specific ID (INSERT OR REPLACE)
        columns = ["id", "date", "serial_no", "card_id", "name", "guardian", "age", "sex", "area", "pregnant",
                   "address", "phone", "ref_from", "disease_case", "symptoms", "paraclinical", "diagnosis",
                   "treatment", "imci", "nutrition", "ref_to", "service", "remark", "patient_type", "branch_code"]
        placeholders = ",".join(["?"] * len(columns))
        query = f"INSERT OR REPLACE INTO patient ({','.join(columns)}) VALUES ({placeholders})"
        params = (id, date, serial_no, card_id, name, guardian, age, sex, area, pregnant, address, phone,
                  ref_from, disease_case, symptoms, paraclinical, diagnosis, treatment,
                  imci, nutrition, ref_to, service, remark, patient_type, branch_code)

        execute_write(query, params)
        return id
    else:
        # Regular insert (auto-increment ID) - use single connection to avoid race condition
        columns = _patient_insert_columns()
        placeholders = ",".join(["?"] * len(columns))
        query = f"INSERT INTO patient ({','.join(columns)}) VALUES ({placeholders})"
        params = (date, serial_no, card_id, name, guardian, age, sex, area, pregnant, address, phone,
                  ref_from, disease_case, symptoms, paraclinical, diagnosis, treatment,
                  imci, nutrition, ref_to, service, remark, patient_type, branch_code)

        # Get lastrowid from the same connection to avoid race condition
        return execute_write(query, params, return_lastrowid=True) or 0

def view(branch_code=None):
    if branch_code:
        branch_clause, branch_params = _branch_filter_clause(branch_code)
        return execute_read(f"SELECT * FROM patient WHERE {branch_clause} ORDER BY id ASC", tuple(branch_params))
    return execute_read("SELECT * FROM patient ORDER BY id ASC")

def view_by_patient_type(patient_type, branch_code=None):
    """
    View patients filtered by patient type (Child or Adult)
    
    Args:
        patient_type: 'Child' or 'Adult'
    
    Returns:
        List of patient records for the specified type
    """
    # Handle cases where patient_type might be NULL or empty in old data
    # Treat NULL/empty as 'Child' for backward compatibility
    clause, params = _build_patient_type_filter(patient_type)
    branch_clause, branch_params = _branch_filter_clause(branch_code) if branch_code else ("1=1", [])
    return execute_read(f"SELECT * FROM patient WHERE {clause} AND {branch_clause} ORDER BY id ASC", tuple(params + branch_params))

def delete(id, branch_code=None):
    if branch_code:
        branch_clause, branch_params = _branch_filter_clause(branch_code)
        execute_write(f"DELETE FROM patient WHERE id=? AND {branch_clause}", (id, *branch_params))
    else:
        execute_write("DELETE FROM patient WHERE id=?", (id,))

def delete_all(branch_code=None):
    if branch_code:
        branch_clause, branch_params = _branch_filter_clause(branch_code)
        execute_write(f"DELETE FROM patient WHERE {branch_clause}", tuple(branch_params))
    else:
        execute_write("DELETE FROM patient")

def update(id, date, serial_no, card_id, name, guardian, age, sex, area, pregnant, address, phone,
           ref_from, disease_case, symptoms, paraclinical, diagnosis, treatment,
           imci, nutrition, ref_to, service, remark, patient_type, branch_code=None):
    """
    Update a patient record with all fields
    
    Args:
        id: Patient ID
        All other args: Patient data fields
    """
    patient_type = _infer_patient_type(serial_no, patient_type, age)
    branch_set = ""
    branch_params = []
    if branch_code:
        branch_set = ", branch_code=?"
        branch_params.append(normalize_branch_code(branch_code))
    execute_write("""
        UPDATE patient SET
        date=?, serial_no=?, card_id=?, name=?, guardian=?, age=?, sex=?, area=?, pregnant=?,
        address=?, phone=?, ref_from=?, disease_case=?, symptoms=?, paraclinical=?,
        diagnosis=?, treatment=?, imci=?, nutrition=?, ref_to=?, service=?, remark=?, patient_type=?{branch_set}
        WHERE id=?
    """.format(branch_set=branch_set), (date, serial_no, card_id, name, guardian, age, sex, area, pregnant, address, phone,
          ref_from, disease_case, symptoms, paraclinical, diagnosis, treatment, imci, nutrition, ref_to, service, remark, patient_type, *branch_params, id))

def update_serial(id, serial_no):
    """
    Update only the serial number for a patient
    
    Args:
        id: Patient ID
        serial_no: New serial number
    """
    execute_write("UPDATE patient SET serial_no=? WHERE id=?", (serial_no, id))

def update_fields(id, fields_dict):
    """
    Update specific fields for a patient (dynamic update)
    
    Args:
        id: Patient ID
        fields_dict: Dictionary of {column_name: value} to update
    
    Returns:
        bool: True if update was successful
    
    Example:
        >>> update_fields(1, {'name': 'New Name', 'phone': '123456'})
    """
    if not fields_dict:
        return False

    invalid_columns = [col for col in fields_dict.keys() if col not in PATIENT_COLUMNS]
    if invalid_columns:
        raise ValueError(f"Invalid patient columns: {', '.join(sorted(invalid_columns))}")

    # Build SET clause dynamically
    set_clause = ", ".join([f"{col}=?" for col in fields_dict.keys()])
    params = list(fields_dict.values()) + [id]
    
    query = f"UPDATE patient SET {set_clause} WHERE id=?"
    execute_write(query, params)
    return True

def update_counters_batch(data_list):
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.executemany("UPDATE patient SET area=?, disease_case=?, imci=? WHERE id=?", data_list)
        conn.commit()

def search(keyword="", patient_type=None, branch_code=None):
    # Search by Name OR Serial No OR Card ID OR Phone
    # Optionally filter by patient type
    if patient_type:
        type_clause, type_params = _build_patient_type_filter(patient_type)
        branch_clause, branch_params = _branch_filter_clause(branch_code) if branch_code else ("1=1", [])
        query = """
            SELECT * FROM patient
            WHERE (name LIKE ? OR serial_no LIKE ? OR card_id LIKE ? OR phone LIKE ?)
            AND {type_clause}
            AND {branch_clause}
            ORDER BY id ASC
        """.format(type_clause=type_clause, branch_clause=branch_clause)
        pattern = '%' + keyword + '%'
        return execute_read(query, (pattern, pattern, pattern, pattern, *type_params, *branch_params))
    else:
        branch_clause, branch_params = _branch_filter_clause(branch_code) if branch_code else ("1=1", [])
        query = """
            SELECT * FROM patient
            WHERE (name LIKE ? OR serial_no LIKE ? OR card_id LIKE ? OR phone LIKE ?)
            AND {branch_clause}
            ORDER BY id ASC
        """.format(branch_clause=branch_clause)
        pattern = '%' + keyword + '%'
        return execute_read(query, (pattern, pattern, pattern, pattern, *branch_params))

def check_user(username, password):
    user = execute_read("SELECT * FROM users WHERE username=?", (username,), one=True)
    if not user:
        return None

    user_dict = dict(user)
    is_valid, needs_upgrade = _verify_password(password, user_dict["password"])
    if not is_valid:
        return None

    if needs_upgrade:
        execute_write(
            "UPDATE users SET password=? WHERE username=?",
            (_hash_password(password), username)
        )
        user = execute_read("SELECT * FROM users WHERE username=?", (username,), one=True)
        return user

    return user

def add_user(username, password, branch_code="MAIN", role="user"):
    hashed_password = _hash_password(password)
    role = str(role or "user").strip().lower()
    if role not in ("admin", "user"):
        role = "user"
    branch_code = "ALL" if role == "admin" else normalize_branch_code(branch_code)
    execute_write(
        "INSERT INTO users (username, password, role, branch_code) VALUES (?, ?, ?, ?)",
        (username, hashed_password, role, branch_code)
    )

def check_username(username):
    return execute_read("SELECT * FROM users WHERE username=?", (username,), one=True)

def log_login(username):
    login_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    execute_write("INSERT INTO login_history VALUES (NULL, ?, ?)", (username, login_time))

def change_password(username, old_password, new_password):
    """
    Change user password after validating old password

    Args:
        username: Username to change password for
        old_password: Current password to validate
        new_password: New password to set

    Returns:
        tuple: (success: bool, message: str)
    """
    # Verify old password
    if not check_user(username, old_password):
        return False, "Old password is incorrect"

    # Validate new password is not empty
    if not new_password or not new_password.strip():
        return False, "New password cannot be empty"

    # Validate new password is different from old
    if new_password == old_password:
        return False, "New password must be different from old password"

    # Update password
    hashed_password = _hash_password(new_password)
    execute_write("UPDATE users SET password=? WHERE username=?", (hashed_password, username))
    return True, "Password changed successfully"

def get_max_counter_for_category(column_name, category, patient_type=None, branch_code=None):
    # White-list the allowed columns to prevent potential SQL injection
    allowed_columns = ["area", "disease_case", "imci"]
    if column_name not in allowed_columns:
        return 0

    # Build query with optional patient_type filter
    query = f"""
        SELECT COALESCE(MAX(CAST(SUBSTR({column_name}, INSTR({column_name}, '::') + 2) AS INTEGER)), 0)
        FROM patient
        WHERE {column_name} LIKE ?
    """
    
    # Add patient_type filter if provided
    if patient_type:
        type_clause, type_params = _build_patient_type_filter(patient_type)
        query += f" AND {type_clause}"
        logger.debug(f"Query: {query}")
        logger.debug(f"Params: category='{category}', patient_type='{patient_type}'")
        result = execute_read(query, (f'{category}::%', *type_params), one=True)
    else:
        logger.debug(f"Query (no patient_type): {query}")
        logger.debug(f"Params: category='{category}'")
        result = execute_read(query, (f'{category}::%',), one=True)

    if branch_code:
        branch_clause, branch_params = _branch_filter_clause(branch_code)
        query = f"""
            SELECT COALESCE(MAX(CAST(SUBSTR({column_name}, INSTR({column_name}, '::') + 2) AS INTEGER)), 0)
            FROM patient
            WHERE {column_name} LIKE ?
        """
        params = [f'{category}::%']
        if patient_type:
            type_clause, type_params = _build_patient_type_filter(patient_type)
            query += f" AND {type_clause}"
            params.extend(type_params)
        query += f" AND {branch_clause}"
        params.extend(branch_params)
        result = execute_read(query, tuple(params), one=True)

    logger.debug(f"Result: {result}")
    return result[0] if result else 0

def get_max_serial_for_type(patient_type, date_text=None, branch_code=None):
    date_filter = ""
    date_params = ()
    if date_text:
        month_part = _normalize_merge_text(date_text)[3:10]
        if len(month_part) == 7:
            date_filter = " AND SUBSTR(date, 4, 7) = ?"
            date_params = (month_part,)

    if patient_type == 'Child':
        query = """
            SELECT COALESCE(MAX(
                CASE
                    WHEN serial_no LIKE 'C-%' THEN CAST(SUBSTR(serial_no, 3) AS INTEGER)
                    WHEN patient_type IS NULL OR patient_type = '' THEN CAST(serial_no AS INTEGER)
                    WHEN patient_type = 'Child' THEN CAST(serial_no AS INTEGER)
                    ELSE 0
                END
            ), 0)
            FROM patient
            WHERE (patient_type = 'Child' OR patient_type IS NULL OR patient_type = '')
        """
        if date_filter:
            query = f"""
                SELECT COALESCE(MAX(
                    CASE
                        WHEN serial_no LIKE 'C-%' THEN CAST(SUBSTR(serial_no, 3) AS INTEGER)
                        WHEN patient_type IS NULL OR patient_type = '' THEN CAST(serial_no AS INTEGER)
                        WHEN patient_type = 'Child' THEN CAST(serial_no AS INTEGER)
                        ELSE 0
                    END
                ), 0)
                FROM patient
                WHERE (patient_type = 'Child' OR patient_type IS NULL OR patient_type = ''){date_filter}
            """
        params = list(date_params)
    else:
        query = """
            SELECT COALESCE(MAX(
                CASE
                    WHEN serial_no LIKE 'A-%' THEN CAST(SUBSTR(serial_no, 3) AS INTEGER)
                    ELSE CAST(serial_no AS INTEGER)
                END
            ), 0)
            FROM patient
            WHERE patient_type = 'Adult'
        """
        if date_filter:
            query = f"""
                SELECT COALESCE(MAX(
                    CASE
                        WHEN serial_no LIKE 'A-%' THEN CAST(SUBSTR(serial_no, 3) AS INTEGER)
                        ELSE CAST(serial_no AS INTEGER)
                    END
                ), 0)
                FROM patient
                WHERE patient_type = 'Adult'{date_filter}
            """
        params = list(date_params)
    if branch_code:
        branch_clause, branch_params = _branch_filter_clause(branch_code)
        query += f" AND {branch_clause}"
        params.extend(branch_params)
    result = execute_read(query, tuple(params), one=True)
    return result[0] if result else 0

def get_statistics(branch_code=None):
    today_str = datetime.now().strftime("%d/%m/%Y")
    branch_clause, branch_params = _branch_filter_clause(branch_code) if branch_code else ("1=1", [])
    
    def get_stats_for_type(p_type):
        """Helper function to get statistics for a specific patient type"""
        query = """
        SELECT
            COUNT(*) as total,
            SUM(CASE WHEN date = ? THEN 1 ELSE 0 END) as today_count,
            SUM(CASE WHEN sex = 'ប្រុស' THEN 1 ELSE 0 END) as male,
            SUM(CASE WHEN sex = 'ស្រី' THEN 1 ELSE 0 END) as female,

            -- ករណី ថ្មី/ចាស់
            SUM(CASE WHEN disease_case LIKE 'ថ្មី%' THEN 1 ELSE 0 END) as new_total,
            SUM(CASE WHEN disease_case LIKE 'ថ្មី%' AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as new_m,
            SUM(CASE WHEN disease_case LIKE 'ថ្មី%' AND sex = 'ស្រី' THEN 1 ELSE 0 END) as new_f,
            SUM(CASE WHEN disease_case LIKE 'ចាស់%' THEN 1 ELSE 0 END) as old_total,
            SUM(CASE WHEN disease_case LIKE 'ចាស់%' AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as old_m,
            SUM(CASE WHEN disease_case LIKE 'ចាស់%' AND sex = 'ស្រី' THEN 1 ELSE 0 END) as old_f,

            -- តំបន់ ក
            SUM(CASE WHEN area LIKE 'ក%' AND disease_case LIKE 'ថ្មី%' THEN 1 ELSE 0 END) as area_a_new,
            SUM(CASE WHEN area LIKE 'ក%' AND disease_case LIKE 'ថ្មី%' AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as area_a_new_m,
            SUM(CASE WHEN area LIKE 'ក%' AND disease_case LIKE 'ថ្មី%' AND sex = 'ស្រី' THEN 1 ELSE 0 END) as area_a_new_f,
            SUM(CASE WHEN area LIKE 'ក%' AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as area_a_total_m,
            SUM(CASE WHEN area LIKE 'ក%' AND sex = 'ស្រី' THEN 1 ELSE 0 END) as area_a_total_f,

            -- តំបន់ ខ
            SUM(CASE WHEN area LIKE 'ខ%' AND disease_case LIKE 'ថ្មី%' THEN 1 ELSE 0 END) as area_b_new,
            SUM(CASE WHEN area LIKE 'ខ%' AND disease_case LIKE 'ថ្មី%' AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as area_b_new_m,
            SUM(CASE WHEN area LIKE 'ខ%' AND disease_case LIKE 'ថ្មី%' AND sex = 'ស្រី' THEN 1 ELSE 0 END) as area_b_new_f,
            SUM(CASE WHEN area LIKE 'ខ%' AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as area_b_total_m,
            SUM(CASE WHEN area LIKE 'ខ%' AND sex = 'ស្រី' THEN 1 ELSE 0 END) as area_b_total_f,

            -- តំបន់ គ
            SUM(CASE WHEN area LIKE 'គ%' AND disease_case LIKE 'ថ្មី%' THEN 1 ELSE 0 END) as area_c_new,
            SUM(CASE WHEN area LIKE 'គ%' AND disease_case LIKE 'ថ្មី%' AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as area_c_new_m,
            SUM(CASE WHEN area LIKE 'គ%' AND disease_case LIKE 'ថ្មី%' AND sex = 'ស្រី' THEN 1 ELSE 0 END) as area_c_new_f,
            SUM(CASE WHEN area LIKE 'គ%' AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as area_c_total_m,
            SUM(CASE WHEN area LIKE 'គ%' AND sex = 'ស្រី' THEN 1 ELSE 0 END) as area_c_total_f,

            -- សេវាកម្ម (Services)
            SUM(CASE WHEN service_norm IN ('PAY', 'P') AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as pay_m,
            SUM(CASE WHEN service_norm IN ('PAY', 'P') AND sex = 'ស្រី' THEN 1 ELSE 0 END) as pay_f,
            SUM(CASE WHEN service_norm = 'HEF' AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as hef_m,
            SUM(CASE WHEN service_norm = 'HEF' AND sex = 'ស្រី' THEN 1 ELSE 0 END) as hef_f,
            SUM(CASE WHEN service_norm IN ('FREE', 'E') AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as free_m,
            SUM(CASE WHEN service_norm IN ('FREE', 'E') AND sex = 'ស្រី' THEN 1 ELSE 0 END) as free_f,
            SUM(CASE WHEN service_norm IN ('HEFR', 'HEF-R') AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as hefr_m,
            SUM(CASE WHEN service_norm IN ('HEFR', 'HEF-R') AND sex = 'ស្រី' THEN 1 ELSE 0 END) as hefr_f,
            SUM(CASE WHEN service_norm IN ('HEFI', 'HEF-I') AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as hefi_m,
            SUM(CASE WHEN service_norm IN ('HEFI', 'HEF-I') AND sex = 'ស្រី' THEN 1 ELSE 0 END) as hefi_f,
            SUM(CASE WHEN service_norm IN ('NSSF-8', 'បសស-8', 'NSSF', 'បសស') AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as nssf8_m,
            SUM(CASE WHEN service_norm IN ('NSSF-8', 'បសស-8', 'NSSF', 'បសស') AND sex = 'ស្រី' THEN 1 ELSE 0 END) as nssf8_f,
            SUM(CASE WHEN (service_norm = 'NSSF-7' OR service_norm = 'បសស-7') AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as nssf7_m,
            SUM(CASE WHEN (service_norm = 'NSSF-7' OR service_norm = 'បសស-7') AND sex = 'ស្រី' THEN 1 ELSE 0 END) as nssf7_f,
            SUM(CASE WHEN (service_norm = 'NSSF-A' OR service_norm IN ('បសស-ឃ', 'បសស-ឌ', 'បសស-ធ', 'បសស-ស')) AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as nssfa_m,
            SUM(CASE WHEN (service_norm = 'NSSF-A' OR service_norm IN ('បសស-ឃ', 'បសស-ឌ', 'បសស-ធ', 'បសស-ស')) AND sex = 'ស្រី' THEN 1 ELSE 0 END) as nssfa_f,
            SUM(CASE WHEN (
                service_norm = 'OTHER'
                OR service_norm = ''
                OR service_norm NOT IN ('PAY', 'P', 'HEF', 'FREE', 'E', 'HEFR', 'HEF-R', 'HEFI', 'HEF-I', 'NSSF-8', 'បសស-8', 'NSSF-7', 'បសស-7', 'NSSF', 'NSSF-A', 'បសស', 'បសស-ឃ', 'បសស-ឌ', 'បសស-ធ', 'បសស-ស')
            ) AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as other_m,
            SUM(CASE WHEN (
                service_norm = 'OTHER'
                OR service_norm = ''
                OR service_norm NOT IN ('PAY', 'P', 'HEF', 'FREE', 'E', 'HEFR', 'HEF-R', 'HEFI', 'HEF-I', 'NSSF-8', 'បសស-8', 'NSSF-7', 'បសស-7', 'NSSF', 'NSSF-A', 'បសស', 'បសស-ឃ', 'បសស-ឌ', 'បសស-ធ', 'បសស-ស')
            ) AND sex = 'ស្រី' THEN 1 ELSE 0 END) as other_f,

            -- អាយុកុមារ (Child Age Groups) - គណនាពី age column
            SUM(CASE WHEN 
                (age LIKE '០-២៩ថ្ងៃ%' OR age LIKE '0-28%' OR age LIKE '0-29%')
                AND patient_type = 'Child'
            THEN 1 ELSE 0 END) as age_0_29_days,
            SUM(CASE WHEN 
                (age LIKE '២៩ថ្ងៃ-១១ខែ%' OR age LIKE '29-11%' OR age LIKE '29ថ្ងៃ-11ខែ%')
                AND patient_type = 'Child'
            THEN 1 ELSE 0 END) as age_29days_11months,
            SUM(CASE WHEN 
                (age LIKE '១-៤ឆ្នាំ%' OR age LIKE '1-4%')
                AND patient_type = 'Child'
            THEN 1 ELSE 0 END) as age_1_4_years
        FROM (
            SELECT
                *,
                UPPER(REPLACE(REPLACE(TRIM(COALESCE(service, '')), '\\', ''), ' ', '-')) as service_norm
            FROM patient
        )
        WHERE patient_type = ? AND {branch_clause}
        """
        query = query.format(branch_clause=branch_clause)
        row = execute_read(query, (today_str, p_type, *branch_params), one=True)
        keys = [
            'total', 'today_count', 'male', 'female',
            'new_total', 'new_m', 'new_f', 'old_total', 'old_m', 'old_f',
            'area_a_new', 'area_a_new_m', 'area_a_new_f', 'area_a_total_m', 'area_a_total_f',
            'area_b_new', 'area_b_new_m', 'area_b_new_f', 'area_b_total_m', 'area_b_total_f',
            'area_c_new', 'area_c_new_m', 'area_c_new_f', 'area_c_total_m', 'area_c_total_f',
            'pay_m', 'pay_f', 'hef_m', 'hef_f',
            'free_m', 'free_f', 'hefr_m', 'hefr_f', 'hefi_m', 'hefi_f',
            'nssf8_m', 'nssf8_f', 'nssf7_m', 'nssf7_f', 'nssfa_m', 'nssfa_f',
            'other_m', 'other_f',
            'age_0_29_days', 'age_29days_11months', 'age_1_4_years'
        ]

        if row:
            processed_row = [r if r is not None else 0 for r in row]
        else:
            processed_row = [0] * len(keys)

        return dict(zip(keys, processed_row))
    
    # Get overall statistics (no patient_type filter)
    overall_query = """
    SELECT
        COUNT(*) as total,
        SUM(CASE WHEN date = ? THEN 1 ELSE 0 END) as today_count,
        SUM(CASE WHEN sex = 'ប្រុស' THEN 1 ELSE 0 END) as male,
        SUM(CASE WHEN sex = 'ស្រី' THEN 1 ELSE 0 END) as female
    FROM patient
    WHERE {branch_clause}
    """
    overall_query = overall_query.format(branch_clause=branch_clause)
    overall_row = execute_read(overall_query, (today_str, *branch_params), one=True)
    
    stats = {
        'child': get_stats_for_type('Child'),
        'adult': get_stats_for_type('Adult'),
        'overall': {
            'total': overall_row[0] if overall_row and overall_row[0] else 0,
            'today_count': overall_row[1] if overall_row and overall_row[1] else 0,
            'male': overall_row[2] if overall_row and overall_row[2] else 0,
            'female': overall_row[3] if overall_row and overall_row[3] else 0,
        }
    }
    
    return stats

def get_diagnosis_statistics(new_only=False):
    """Return patient totals grouped by diagnosis and age range."""
    where_clause = "WHERE disease_case LIKE 'ថ្មី%'" if new_only else ""
    query = """
    SELECT
        COALESCE(NULLIF(TRIM(diagnosis), ''), 'មិនបានបញ្ជាក់') as diagnosis_name,
        COUNT(*) as total,
        SUM(CASE WHEN patient_type = 'Child' OR patient_type IS NULL OR patient_type = '' THEN 1 ELSE 0 END) as child_total,
        SUM(CASE WHEN patient_type = 'Adult' THEN 1 ELSE 0 END) as adult_total,
        SUM(CASE WHEN sex = 'ប្រុស' THEN 1 ELSE 0 END) as male,
        SUM(CASE WHEN sex = 'ស្រី' THEN 1 ELSE 0 END) as female,
        SUM(CASE WHEN TRIM(COALESCE(ref_to, '')) != '' THEN 1 ELSE 0 END) as ref_to_total,
        SUM(CASE WHEN UPPER(TRIM(service)) IN ('HEF', 'HEFR', 'HEF-R', 'HEFI', 'HEF-I') THEN 1 ELSE 0 END) as hef_total,
        SUM(CASE WHEN (age LIKE '០-២៩ថ្ងៃ%' OR age LIKE '0-28%' OR age LIKE '0-29%') AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as age_0_29_days_m,
        SUM(CASE WHEN (age LIKE '០-២៩ថ្ងៃ%' OR age LIKE '0-28%' OR age LIKE '0-29%') AND sex = 'ស្រី' THEN 1 ELSE 0 END) as age_0_29_days_f,
        SUM(CASE WHEN (age LIKE '២៩ថ្ងៃ-១១ខែ%' OR age LIKE '29-11%' OR age LIKE '29ថ្ងៃ-11ខែ%' OR age LIKE '29ថ្ងៃ-១១ខែ%') AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as age_29days_11months_m,
        SUM(CASE WHEN (age LIKE '២៩ថ្ងៃ-១១ខែ%' OR age LIKE '29-11%' OR age LIKE '29ថ្ងៃ-11ខែ%' OR age LIKE '29ថ្ងៃ-១១ខែ%') AND sex = 'ស្រី' THEN 1 ELSE 0 END) as age_29days_11months_f,
        SUM(CASE WHEN (age LIKE '១-៤ឆ្នាំ%' OR age LIKE '1-4%') AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as age_1_4_years_m,
        SUM(CASE WHEN (age LIKE '១-៤ឆ្នាំ%' OR age LIKE '1-4%') AND sex = 'ស្រី' THEN 1 ELSE 0 END) as age_1_4_years_f,
        SUM(CASE WHEN (age LIKE '៥-១៤ឆ្នាំ%' OR age LIKE '5-14%') AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as age_5_14_years_m,
        SUM(CASE WHEN (age LIKE '៥-១៤ឆ្នាំ%' OR age LIKE '5-14%') AND sex = 'ស្រី' THEN 1 ELSE 0 END) as age_5_14_years_f,
        SUM(CASE WHEN (age LIKE '១៥-២៤ឆ្នាំ%' OR age LIKE '15-24%') AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as age_15_24_years_m,
        SUM(CASE WHEN (age LIKE '១៥-២៤ឆ្នាំ%' OR age LIKE '15-24%') AND sex = 'ស្រី' THEN 1 ELSE 0 END) as age_15_24_years_f,
        SUM(CASE WHEN (age LIKE '២៥-៤៩ឆ្នាំ%' OR age LIKE '25-49%') AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as age_25_49_years_m,
        SUM(CASE WHEN (age LIKE '២៥-៤៩ឆ្នាំ%' OR age LIKE '25-49%') AND sex = 'ស្រី' THEN 1 ELSE 0 END) as age_25_49_years_f,
        SUM(CASE WHEN (age LIKE '៥០-៦៤ឆ្នាំ%' OR age LIKE '50-64%') AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as age_50_64_years_m,
        SUM(CASE WHEN (age LIKE '៥០-៦៤ឆ្នាំ%' OR age LIKE '50-64%') AND sex = 'ស្រី' THEN 1 ELSE 0 END) as age_50_64_years_f,
        SUM(CASE WHEN (age LIKE '>=៦៥ឆ្នាំ%' OR age LIKE '>៦៥ឆ្នាំ%' OR age LIKE '>=៦៤ឆ្នាំ%' OR age LIKE '>៦៤ឆ្នាំ%' OR age LIKE '>=65%' OR age LIKE '>65%' OR age LIKE '>=64%' OR age LIKE '>64%' OR age LIKE '65%') AND sex = 'ប្រុស' THEN 1 ELSE 0 END) as age_64_plus_m,
        SUM(CASE WHEN (age LIKE '>=៦៥ឆ្នាំ%' OR age LIKE '>៦៥ឆ្នាំ%' OR age LIKE '>=៦៤ឆ្នាំ%' OR age LIKE '>៦៤ឆ្នាំ%' OR age LIKE '>=65%' OR age LIKE '>65%' OR age LIKE '>=64%' OR age LIKE '>64%' OR age LIKE '65%') AND sex = 'ស្រី' THEN 1 ELSE 0 END) as age_64_plus_f
    FROM patient
    {where_clause}
    GROUP BY diagnosis_name
    ORDER BY total DESC, diagnosis_name ASC
    """.format(where_clause=where_clause)
    rows = execute_read(query) or []
    return [
        {
            "diagnosis": row["diagnosis_name"],
            "total": row["total"] or 0,
            "child": row["child_total"] or 0,
            "adult": row["adult_total"] or 0,
            "male": row["male"] or 0,
            "female": row["female"] or 0,
            "ref_to_total": row["ref_to_total"] or 0,
            "hef_total": row["hef_total"] or 0,
            "age_0_29_days_m": row["age_0_29_days_m"] or 0,
            "age_0_29_days_f": row["age_0_29_days_f"] or 0,
            "age_29days_11months_m": row["age_29days_11months_m"] or 0,
            "age_29days_11months_f": row["age_29days_11months_f"] or 0,
            "age_1_4_years_m": row["age_1_4_years_m"] or 0,
            "age_1_4_years_f": row["age_1_4_years_f"] or 0,
            "age_5_14_years_m": row["age_5_14_years_m"] or 0,
            "age_5_14_years_f": row["age_5_14_years_f"] or 0,
            "age_15_24_years_m": row["age_15_24_years_m"] or 0,
            "age_15_24_years_f": row["age_15_24_years_f"] or 0,
            "age_25_49_years_m": row["age_25_49_years_m"] or 0,
            "age_25_49_years_f": row["age_25_49_years_f"] or 0,
            "age_50_64_years_m": row["age_50_64_years_m"] or 0,
            "age_50_64_years_f": row["age_50_64_years_f"] or 0,
            "age_64_plus_m": row["age_64_plus_m"] or 0,
            "age_64_plus_f": row["age_64_plus_f"] or 0,
        }
        for row in rows
    ]

def get_login_history():
    return execute_read("SELECT username, login_time FROM login_history ORDER BY id DESC")

def close_connections():
    """
    Close any active database connections
    This is useful before copying/moving the database file
    """
    # SQLite connections are automatically closed when using 'with' statement
    # This function is a placeholder for future connection pool management
    pass

def advanced_search(start_date, end_date, sex, disease, area, p_type, branch_code=None):
    """
    Performs an efficient, multi-criteria search directly in the database.
    """
    # Base query
    query_parts = ["SELECT * FROM patient WHERE 1=1"]
    params = []

    # Date filtering using SQL SUBSTR to format DD/MM/YYYY into YYYYMMDD for comparison
    if start_date and end_date:
        start_date_sql = start_date.strftime('%Y%m%d')
        end_date_sql = end_date.strftime('%Y%m%d')
        query_parts.append("AND (SUBSTR(date, 7, 4) || SUBSTR(date, 4, 2) || SUBSTR(date, 1, 2)) BETWEEN ? AND ?")
        params.extend([start_date_sql, end_date_sql])

    # Text field filtering
    if sex != "All":
        query_parts.append("AND sex = ?")
        params.append(sex)

    if p_type != "All":
        type_clause, type_params = _build_patient_type_filter(p_type)
        query_parts.append(f"AND {type_clause}")
        params.extend(type_params)

    # Composite field filtering (using LIKE)
    if disease != "All":
        query_parts.append("AND disease_case LIKE ?")
        params.append(f"{disease}%")

    if area != "All":
        query_parts.append("AND area LIKE ?")
        params.append(f"{area}%")

    if branch_code:
        branch_clause, branch_params = _branch_filter_clause(branch_code)
        query_parts.append(f"AND {branch_clause}")
        params.extend(branch_params)

    # Final query
    query = " ".join(query_parts)
    query += " ORDER BY id ASC"
    
    return execute_read(query, tuple(params))

connect()
