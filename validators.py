import re
from datetime import datetime


def validate_date(date_str: str) -> bool:
    """Enforces DD.MM.YYYY date format."""
    if not date_str:
        return False
    try:
        datetime.strptime(date_str.strip(), "%d.%m.%Y")
        return True
    except ValueError:
        return False


def validate_quote_number(quote_str: str) -> bool:
    """Enforces format: SCT - XXXX (e.g., SCT - 7725)."""
    if not quote_str:
        return False
    pattern = r"^SCT\s*-\s*\d+$"
    return bool(re.match(pattern, quote_str.strip(), re.IGNORECASE))


def validate_iob_number(iob_str: str) -> bool:
    """Enforces format: IOB XXXXX (e.g., IOB 22941)."""
    if not iob_str:
        return False
    pattern = r"^IOB\s+\d+$"
    return bool(re.match(pattern, iob_str.strip(), re.IGNORECASE))


def validate_required_fields(field_dict: dict) -> list:
    """Checks that no critical field defaults to 'N/A' or empty silently."""
    critical_fields = [
        "client_name",
        "document_title",
        "job_number",
        "part_description",
        "quantity",
        "responsible_person",
        "customer",
        "quote_number",
        "surcotec_ref_number",
        "date_created",
        "due_date",
        "qcp_steps",
    ]
    missing_fields = []
    for field in critical_fields:
        val = str(field_dict.get(field, "")).strip()
        if not val or val.upper() == "N/A":
            missing_fields.append(field)
    return missing_fields