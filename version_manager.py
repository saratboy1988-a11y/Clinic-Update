# -*- coding: utf-8 -*-
"""
Version Management Module
Centralized version information for the Clinic Management System
"""

import json
import os

# ============================================================================
# CENTRALIZED VERSION DEFINITION
# ============================================================================
# Update this version number here ONLY when releasing a new version
# This file is the SINGLE SOURCE OF TRUTH for version numbers
# ============================================================================

APP_VERSION = "2.1.2"  # <-- UPDATE THIS LINE for new releases

# Version format: MAJOR.MINOR.PATCH
# - MAJOR: Breaking changes or major new features
# - MINOR: New features (backward compatible)
# - PATCH: Bug fixes and minor improvements


def get_version():
    """Get the current application version"""
    return APP_VERSION


def get_version_info():
    """Get detailed version information"""
    return {
        "version": APP_VERSION,
        "build_date": __import__('datetime').datetime.now().strftime("%Y-%m-%d"),
        "copyright": "Copyright © 2026 NOU SARAT",
        "author": "នូរ សារ៉ាត់ (NOU SARAT)"
    }


def update_version_file(version_file=None):
    """
    Update version.json with the current version
    
    Args:
        version_file: Path to version.json (defaults to version.json in script directory)
    
    Returns:
        bool: True if successful, False otherwise
    """
    if version_file is None:
        version_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "version.json")
    
    try:
        # Read existing version.json to preserve other fields
        if os.path.exists(version_file):
            with open(version_file, 'r', encoding='utf-8') as f:
                version_data = json.load(f)
        else:
            version_data = {}
        
        # Update version
        version_data["version"] = APP_VERSION
        
        # Write back
        with open(version_file, 'w', encoding='utf-8') as f:
            json.dump(version_data, f, indent=2, ensure_ascii=False)
        
        return True
    except Exception as e:
        print(f"Error updating version file: {e}")
        return False


def generate_iss_version_line():
    """
    Generate the Inno Setup Script version definition line
    
    Returns:
        str: The #define line for .iss file
    """
    return f'#define MyAppVersion "{APP_VERSION}"'
