"""
License Manager Module
Handles license key generation and validation for Clinic Management System.

Author: នូរ សារ៉ាត់ (NOU SARAT)
"""

import os
import hashlib
import hmac
import base64
from datetime import datetime, timedelta

LICENSE_VERSION = "v2"
DEV_SECRET_SALT = "DEV_ONLY_CHANGE_ME"


def _is_production():
    return os.getenv("CLINIC_ENV", "").strip().lower() == "production"


def _truthy_env(name):
    return os.getenv(name, "").strip().lower() in ("1", "true", "yes", "on")


def _get_secret_salt(require_for_generation=False):
    secret = os.getenv("CLINIC_SECRET_SALT", "").strip()
    if secret:
        return secret
    if _is_production() or require_for_generation:
        raise RuntimeError("CLINIC_SECRET_SALT must be set before using offline licenses.")
    return DEV_SECRET_SALT


def _legacy_license_validation_enabled():
    if _is_production():
        return _truthy_env("CLINIC_ALLOW_LEGACY_LICENSES")
    return True


def _hmac_signature(machine_id, email, expiry, secret):
    payload = f"{machine_id}|{email}|{expiry}".encode("utf-8")
    return hmac.new(secret.encode("utf-8"), payload, hashlib.sha256).hexdigest()


def generate_license_key(email: str, machine_id: str, duration_type: str) -> str:
    """
    Generate a license key based on email, machine ID, and duration type.
    
    Args:
        email: User's email address
        machine_id: Unique machine identifier
        duration_type: License duration (e.g., "1 ខែ", "Lifetime (មួយជីវិត)")
    
    Returns:
        Base64-encoded license key string
    """
    if duration_type == "Lifetime (មួយជីវិត)":
        expiry = "Lifetime"
    else:
        try:
            months = int(duration_type.split(' ')[0])
            expiry_date = datetime.now() + timedelta(days=30 * months)
            expiry = expiry_date.strftime("%Y-%m-%d")
        except (ValueError, IndexError):
            raise ValueError(f"Invalid duration type: {duration_type}")

    # Create a full HMAC-SHA256 signature to prevent tampering.
    secret = _get_secret_salt(require_for_generation=True)
    signature = _hmac_signature(machine_id, email, expiry, secret)

    raw_key = f"{LICENSE_VERSION}|{machine_id}|{email}|{expiry}|{signature}"
    return base64.b64encode(raw_key.encode()).decode()


def validate_license(email: str, machine_id: str, key: str) -> tuple:
    """
    Validate a license key for authenticity and expiry.
    
    Args:
        email: User's email address
        machine_id: Machine identifier to validate against
        key: License key to validate
    
    Returns:
        Tuple of (is_valid: bool, message: str)
        - If valid: (True, expiry_date or "Lifetime")
        - If invalid: (False, error_message)
    """
    try:
        decoded = base64.b64decode(key.encode()).decode()
        parts = decoded.split('|')
        
        if len(parts) == 5 and parts[0] == LICENSE_VERSION:
            _, m_id, mail, expiry, signature = parts
            secret = _get_secret_salt()
            expected_sig = _hmac_signature(m_id, mail, expiry, secret)
        elif len(parts) == 4 and _legacy_license_validation_enabled():
            m_id, mail, expiry, signature = parts
            legacy_secret = os.getenv("CLINIC_LEGACY_SECRET_SALT", "").strip() or _get_secret_salt()
            legacy_check_str = f"{m_id}|{mail}|{expiry}|{legacy_secret}"
            expected_sig = hashlib.sha256(legacy_check_str.encode()).hexdigest()[:10]
        elif len(parts) == 4:
            return False, "Legacy offline licenses are disabled"
        else:
            return False, "Invalid Format"
        
        if m_id != machine_id or mail != email:
            return False, "Wrong PC or Email"
        
        if not hmac.compare_digest(signature, expected_sig):
            return False, "Tampered Key"

        if expiry == "Lifetime":
            return True, "Lifetime"

        exp_date = datetime.strptime(expiry, "%Y-%m-%d")
        if datetime.now() > exp_date:
            return False, "Expired"

        return True, expiry
    
    except Exception as e:
        return False, f"Error: {str(e)}"


def get_license_duration_info(key: str, email: str, machine_id: str) -> dict:
    """
    Get detailed information about a license.
    
    Args:
        key: License key to analyze
        email: Associated email
        machine_id: Associated machine ID
    
    Returns:
        Dictionary with license details or error information
    """
    is_valid, message = validate_license(email, machine_id, key)
    
    if not is_valid:
        return {
            "valid": False,
            "error": message,
            "expiry": None,
            "days_remaining": None
        }
    
    try:
        decoded = base64.b64decode(key.encode()).decode()
        parts = decoded.split('|')
        expiry = parts[3] if len(parts) == 5 and parts[0] == LICENSE_VERSION else parts[2]
        
        if expiry == "Lifetime":
            return {
                "valid": True,
                "type": "Lifetime",
                "expiry": "Lifetime",
                "days_remaining": None
            }
        
        exp_date = datetime.strptime(expiry, "%Y-%m-%d")
        days_remaining = (exp_date - datetime.now()).days
        
        return {
            "valid": True,
            "type": "Time-limited",
            "expiry": expiry,
            "days_remaining": days_remaining
        }
    except Exception as e:
        return {
            "valid": False,
            "error": f"Parse error: {str(e)}",
            "expiry": None,
            "days_remaining": None
        }
