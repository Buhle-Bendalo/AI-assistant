import json
import os
from datetime import datetime

AUDIT_LOG_FILE = "audit.log"


def log_session(
    session_id: str,
    uploaded_filename: str,
    extracted_fields: dict,
    user_corrections: list,
    output_filename: str,
):
    """Logs session activity to audit.log in JSON-lines format for POPIA compliance and traceability."""
    log_entry = {
        "timestamp": datetime.now().isoformat(),
        "session_id": session_id,
        "uploaded_filename": uploaded_filename,
        "extracted_fields": extracted_fields,
        "user_corrections": user_corrections,
        "output_filename": output_filename,
    }

    try:
        with open(AUDIT_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(json.dumps(log_entry) + "\n")
    except Exception as e:
        print(f"Failed to write to audit.log: {str(e)}")