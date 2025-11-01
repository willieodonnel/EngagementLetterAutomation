"""
Data Processing Module for Engagement Letter Automation

This module handles all data collection and processing for engagement letters.
It provides functions to collect loan information, property details, and vendor data,
returning everything in a standardized JSON format.
"""

import json
import os
import pandas as pd
from datetime import datetime, timedelta
from typing import Dict, Optional, Any
import pytz


def collect_engagement_data(
    loan_type: str,
    letter_type: str,
    vendor_info: Optional[Dict[str, str]] = None,
    dates_info: Optional[Dict[str, str]] = None,
    loan_info: Optional[Dict[str, str]] = None,
    property_info: Optional[Dict[str, str]] = None,
    use_autofill: bool = False,
    database_path: Optional[str] = None
) -> Dict[str, Any]:
    """
    Main function to collect all necessary data for an engagement letter.

    Args:
        loan_type: Type of loan ('7a', '504', 'CC')
        letter_type: Type of letter ('App', 'Env', 'Sec', 'SFR', 'Phase 1', 'Phase 2')
        vendor_info: Optional pre-filled vendor information
        dates_info: Optional pre-filled dates information
        loan_info: Optional pre-filled loan information
        property_info: Optional pre-filled property information
        use_autofill: Whether to use autofill features (database lookup, date calculation)
        database_path: Path to vendor database CSV file

    Returns:
        Dictionary containing all collected data in JSON-compatible format
    """

    # Initialize the data structure
    engagement_data = {
        "loan_type": loan_type.upper(),
        "letter_type": letter_type.upper(),
        "vendor": {},
        "dates": {},
        "loan": {},
        "property": {}
    }

    # Collect vendor information
    if vendor_info:
        engagement_data["vendor"] = vendor_info
    elif use_autofill and database_path:
        engagement_data["vendor"] = _autofill_vendor(letter_type, database_path)
    else:
        engagement_data["vendor"] = _collect_vendor_manually()

    # Collect dates information
    if dates_info:
        engagement_data["dates"] = dates_info
    elif use_autofill:
        engagement_data["dates"] = _calculate_dates(letter_type)
    else:
        engagement_data["dates"] = _collect_dates_manually()

    # Collect loan information
    if loan_info:
        engagement_data["loan"] = loan_info
    else:
        engagement_data["loan"] = _collect_loan_info(loan_type, letter_type)

    # Collect property information
    if property_info:
        engagement_data["property"] = property_info
    else:
        engagement_data["property"] = _collect_property_info(letter_type)

    return engagement_data


def collect_dual_engagement_data(
    loan_type: str,
    app_vendor_info: Optional[Dict[str, str]] = None,
    env_vendor_info: Optional[Dict[str, str]] = None,
    app_dates_info: Optional[Dict[str, str]] = None,
    env_dates_info: Optional[Dict[str, str]] = None,
    loan_info: Optional[Dict[str, str]] = None,
    property_info: Optional[Dict[str, str]] = None,
    use_autofill: bool = False,
    database_path: Optional[str] = None
) -> Dict[str, Any]:
    """
    Collect data for dual engagement letters (appraisal and environmental).

    Args:
        loan_type: Type of loan ('7a', '504', 'CC')
        app_vendor_info: Optional pre-filled appraiser vendor information
        env_vendor_info: Optional pre-filled environmental vendor information
        app_dates_info: Optional pre-filled appraisal dates information
        env_dates_info: Optional pre-filled environmental dates information
        loan_info: Optional pre-filled loan information (shared)
        property_info: Optional pre-filled property information (shared)
        use_autofill: Whether to use autofill features
        database_path: Path to vendor database CSV file

    Returns:
        Dictionary containing data for both letters
    """

    dual_data = {
        "loan_type": loan_type.upper(),
        "appraisal": {
            "letter_type": "APP",
            "vendor": {},
            "dates": {}
        },
        "environmental": {
            "letter_type": "ENV",
            "vendor": {},
            "dates": {}
        },
        "shared": {
            "loan": {},
            "property": {}
        }
    }

    # Collect appraisal vendor information
    if app_vendor_info:
        dual_data["appraisal"]["vendor"] = app_vendor_info
    elif use_autofill and database_path:
        print("\n--- Collecting Appraiser Information ---")
        dual_data["appraisal"]["vendor"] = _autofill_vendor("App", database_path)
    else:
        print("\n--- Collecting Appraiser Information ---")
        dual_data["appraisal"]["vendor"] = _collect_vendor_manually()

    # Collect environmental vendor information
    if env_vendor_info:
        dual_data["environmental"]["vendor"] = env_vendor_info
    elif use_autofill and database_path:
        print("\n--- Collecting Environmental Consultant Information ---")
        dual_data["environmental"]["vendor"] = _autofill_vendor("Env", database_path)
    else:
        print("\n--- Collecting Environmental Consultant Information ---")
        dual_data["environmental"]["vendor"] = _collect_vendor_manually()

    # Collect appraisal dates
    if app_dates_info:
        dual_data["appraisal"]["dates"] = app_dates_info
    elif use_autofill:
        print("\n--- Appraisal Delivery Timeline ---")
        dual_data["appraisal"]["dates"] = _calculate_dates("App")
    else:
        print("\n--- Appraisal Delivery Timeline ---")
        dual_data["appraisal"]["dates"] = _collect_dates_manually()

    # Collect environmental dates
    if env_dates_info:
        dual_data["environmental"]["dates"] = env_dates_info
    elif use_autofill:
        print("\n--- Environmental Delivery Timeline ---")
        dual_data["environmental"]["dates"] = _calculate_dates("Env")
    else:
        print("\n--- Environmental Delivery Timeline ---")
        dual_data["environmental"]["dates"] = _collect_dates_manually()

    # Collect shared loan information
    if loan_info:
        dual_data["shared"]["loan"] = loan_info
    else:
        print("\n--- Loan Information ---")
        dual_data["shared"]["loan"] = _collect_loan_info(loan_type, "DUAL")

    # Collect shared property information
    if property_info:
        dual_data["shared"]["property"] = property_info
    else:
        print("\n--- Property Information ---")
        dual_data["shared"]["property"] = _collect_property_info("DUAL")

    # Add fees for each engagement type
    if not app_dates_info or "fee" not in app_dates_info:
        dual_data["appraisal"]["dates"]["fee"] = input("Appraisal fee: $").strip()

    if not env_dates_info or "fee" not in env_dates_info:
        dual_data["environmental"]["dates"]["fee"] = input("Environmental fee: $").strip()

    return dual_data


# Helper functions for autofill features

def _autofill_vendor(vendor_type: str, database_path: str) -> Dict[str, str]:
    """
    Look up vendor information from database with fallback to manual entry.

    Args:
        vendor_type: Type of vendor to look up
        database_path: Path to CSV database file

    Returns:
        Dictionary with vendor information
    """
    try:
        database = pd.read_csv(database_path)

        # Map Phase 1 and Phase 2 to Env for vendor lookup
        lookup_type = vendor_type
        if vendor_type.upper() in ['PHASE 1', 'PHASE 2']:
            lookup_type = 'Env'

        # Filter by vendor type
        filtered = database[database['Type'].str.contains(lookup_type, case=False, na=False)]

        first_name = input(f"Enter vendor's first name (or 'NA' for manual entry): ").strip()

        if first_name.upper() == 'NA':
            return _collect_vendor_manually()

        # Search for vendor
        matches = filtered[filtered['First'].str.contains(first_name, case=False, na=False)]

        if len(matches) == 0:
            print(f"No vendor found with first name '{first_name}'. Switching to manual entry.")
            return _collect_vendor_manually()
        elif len(matches) == 1:
            vendor = matches.iloc[0]
            return {
                "first_name": vendor['First'],
                "last_name": vendor['Last'],
                "company": vendor['Company'],
                "email": vendor['Email'],
                "type": vendor['Type'],
                "region": vendor.get('Region', 'N/A')
            }
        else:
            # Multiple matches, need last name
            last_name = input(f"Multiple matches found. Enter last name: ").strip()
            refined = matches[matches['Last'].str.contains(last_name, case=False, na=False)]

            if len(refined) == 0:
                print(f"No vendor found with name '{first_name} {last_name}'. Switching to manual entry.")
                return _collect_vendor_manually()
            else:
                vendor = refined.iloc[0]
                return {
                    "first_name": vendor['First'],
                    "last_name": vendor['Last'],
                    "company": vendor['Company'],
                    "email": vendor['Email'],
                    "type": vendor['Type'],
                    "region": vendor.get('Region', 'N/A')
                }

    except Exception as e:
        print(f"Error accessing database: {e}. Switching to manual entry.")
        return _collect_vendor_manually()


def _collect_vendor_manually() -> Dict[str, str]:
    """Collect vendor information manually through user input."""
    return {
        "first_name": input("Vendor first name: ").strip(),
        "last_name": input("Vendor last name: ").strip(),
        "company": input("Vendor company: ").strip(),
        "email": input("Vendor email: ").strip(),
        "type": input("Vendor type (App/Env/Sec/SFR): ").strip().upper(),
        "region": input("Vendor region (optional, press Enter to skip): ").strip() or "N/A"
    }


def _calculate_dates(letter_type: str) -> Dict[str, str]:
    """
    Calculate current and delivery dates based on business days.

    Args:
        letter_type: Type of letter (affects default delivery timeline)

    Returns:
        Dictionary with current_date and delivery_date
    """
    # Skip date calculation for secondary appraisals
    if letter_type.upper() == 'SEC':
        pacific = pytz.timezone('America/Los_Angeles')
        current = datetime.now(pacific)
        return {
            "current_date": current.strftime("%#m/%#d/%Y") if os.name == 'nt' else current.strftime("%-m/%-d/%Y"),
            "delivery_date": "N/A"
        }

    # Get delivery timeline
    timeline = input("Delivery timeline (e.g., '10 bds', '2 weeks', '5 days'): ").strip().lower()

    pacific = pytz.timezone('America/Los_Angeles')
    current_date = datetime.now(pacific)
    delivery_date = current_date

    if 'bd' in timeline or 'business' in timeline:
        # Extract number of business days
        days = int(''.join(filter(str.isdigit, timeline)) or '10')

        # Calculate business days (skip weekends)
        business_days = days
        while business_days > 0:
            delivery_date += timedelta(days=1)
            if delivery_date.weekday() < 5:  # Monday = 0, Friday = 4
                business_days -= 1

    elif 'week' in timeline:
        # Extract number of weeks
        weeks = int(''.join(filter(str.isdigit, timeline)) or '1')
        delivery_date += timedelta(weeks=weeks)

    else:
        # Default to calendar days
        days = int(''.join(filter(str.isdigit, timeline)) or '7')
        delivery_date += timedelta(days=days)

    # Use platform-appropriate format for removing leading zeros
    date_format = "%#m/%#d/%Y" if os.name == 'nt' else "%-m/%-d/%Y"

    return {
        "current_date": current_date.strftime(date_format),
        "delivery_date": delivery_date.strftime(date_format)
    }


def _collect_dates_manually() -> Dict[str, str]:
    """Collect date information manually through user input."""
    return {
        "current_date": input("Current date (M/D/YYYY): ").strip(),
        "delivery_date": input("Delivery date (M/D/YYYY or 'N/A'): ").strip()
    }


def _collect_loan_info(loan_type: str, letter_type: str) -> Dict[str, str]:
    """
    Collect loan-specific information.

    Args:
        loan_type: Type of loan ('7a', '504', 'CC')
        letter_type: Type of letter

    Returns:
        Dictionary with loan information
    """
    loan_info = {
        "loan_name": input("Loan name: ").strip(),
        "loan_number": input("Loan number: ").strip()
    }

    # Add CDC company for 504 loans
    if loan_type.upper() == '504':
        loan_info["cdc_company"] = input("CDC company: ").strip()
    else:
        loan_info["cdc_company"] = "N/A"

    return loan_info


def _collect_property_info(letter_type: str) -> Dict[str, str]:
    """
    Collect property-specific information.

    Args:
        letter_type: Type of letter (affects required fields)

    Returns:
        Dictionary with property information
    """
    property_info = {
        "property_address": input("Property address: ").strip(),
        "property_type": input("Property type: ").strip(),
        "sqft": input("Square footage: ").strip(),
        "property_contact_name": input("Property contact name: ").strip(),
        "property_contact_phone": input("Property contact phone: ").strip(),
        "item_to_send": input("Item to send (or 'TBD'): ").strip() or "TBD"
    }

    # Add fee for non-secondary appraisals (if not already collected with dates)
    if letter_type.upper() == 'SEC':
        property_info["fee"] = "600"  # Fixed fee for secondary appraisals
    elif letter_type.upper() != 'DUAL':  # Dual fees are collected separately
        property_info["fee"] = input("Fee: $").strip()

    return property_info


def save_data_to_json(data: Dict[str, Any], filename: str) -> str:
    """
    Save collected data to a JSON file.

    Args:
        data: Dictionary containing all engagement data
        filename: Name of the JSON file to create

    Returns:
        Path to the saved file
    """
    with open(filename, 'w') as f:
        json.dump(data, f, indent=2)
    return filename


def load_data_from_json(filename: str) -> Dict[str, Any]:
    """
    Load engagement data from a JSON file.

    Args:
        filename: Name of the JSON file to load

    Returns:
        Dictionary containing engagement data
    """
    with open(filename, 'r') as f:
        return json.load(f)