import json
import time
from google import genai
from google.genai import types
from tools import (
    generate_excel_workbook,
    read_examples,
    read_Master_template,
    save_example,
)


class SurcotecAgent:

    def __init__(self, api_key: str):
        self.client = genai.Client(api_key=api_key)

        surcotec_tools = [
            read_Master_template,
            generate_excel_workbook,
            read_examples,
            save_example,
        ]

        master_structure = read_Master_template()
        examples_context = read_examples()

        system_instruction = (
            "You are the Surcotec Document Control Assistant. Your job is to read customer "
            "quotation documents and produce a filled-in Excel job pack by calling 'generate_excel_workbook'.\n\n"
            "=== MASTER TEMPLATE STRUCTURE ===\n"
            f"{master_structure}\n\n"
            "=== REFERENCE EXAMPLES ===\n"
            f"{examples_context}\n\n"
            "=== STEP-BY-STEP WORKFLOW ===\n"
            "When a new document is uploaded:\n"
            "1. Call 'read_examples' ONCE to see real input-output mappings.\n"
            "2. Call 'read_Master_template' ONCE to see the template structure.\n"
            "3. Extract ALL required fields from the quotation text.\n"
            "4. Show the user a 'Proposed Change List' with every field value you extracted.\n"
            "5. CRITICAL REQUIREMENT: Every 'Proposed Change List' you generate MUST end with this exact line:\n"
            "   \"Reminder: AI-generated content must be reviewed before finalizing\"\n"
            "6. When the user says 'Produce', 'Save', or 'Generate', call 'generate_excel_workbook' "
            "immediately with all the extracted fields.\n\n"
            "=== FIELD EXTRACTION & VALIDATION RULES ===\n"
            "Extract these exact fields from the quotation:\n\n"
            "client_name: The customer company name (e.g. 'Atlantis Foundries')\n\n"
            "document_title: Construct as 'ST-F-09-01-[IOB_NUMBER] - Rev1 - [CUSTOMER] - [DESCRIPTION]'\n"
            "  Example: 'ST-F-09-01-8111 - Rev1 - Atlantis Foundries - 160O Linear Shafts for HVOF Coating'\n"
            "  The IOB number must be assigned sequentially - check examples to see the last used number.\n\n"
            "job_number: Format MUST follow 'IOB XXXXX' (e.g., 'IOB 22941') - assign next sequential number.\n\n"
            "part_description: The description field from the quote (e.g. '160 O Linear Shafts for HVOF Coating')\n\n"
            "quantity: From the Quantity field in the quote (e.g. '2-off', '3-Off Parts')\n\n"
            "responsible_person: The person who signed the quotation (e.g. 'Sheldon Deysel', 'Ian Walsh')\n\n"
            "customer: The company name from the quote header (e.g. 'Atlantis Foundries')\n\n"
            "quote_number: Format MUST follow 'SCT - XXXX' (e.g. 'SCT - 7725' from 'SCT 7725')\n\n"
            "surcotec_ref_number: The Surcotec Ref. Number from the quote (e.g. format YY-MM-XXX).\n\n"
            "date_created: Format MUST follow DD.MM.YYYY (e.g., '20.01.2026')\n\n"
            "due_date: From the DELIVERY section of the quote (e.g. '10-15 Working Days')\n\n"
            "customer_ref_number: Client Ref. number from quote, or 'N/A' if blank\n\n"
            "customer_job_number: Customer job number if present, else 'N/A'\n\n"
            "customer_drawing_number: Customer drawing number if mentioned, else 'N/A'\n\n"
            "=== QCP STEPS FORMAT ===\n"
            "qcp_steps: A list of process steps from the SCOPE OF WORK, one per line, as:\n"
            "  Step Name|Step description (1-3 sentences of technical detail)\n\n"
            "=== SPRAY MATERIALS FORMAT ===\n"
            "spray_materials: One line per material as 'System|Material Name|Code'\n\n"
            "=== IMPORTANT RULES ===\n"
            "- NEVER call 'generate_excel_workbook' until the user says 'Produce', 'Save', or 'Generate'.\n"
            "- ALWAYS show the Proposed Change List first and wait for user confirmation.\n"
            "- ALWAYS append 'Reminder: AI-generated content must be reviewed before finalizing' at the bottom of every Proposed Change List.\n"
            "- Keep ALL template formatting - only update cell values.\n"
        )

        self.config = types.GenerateContentConfig(
            system_instruction=system_instruction,
            tools=surcotec_tools,
            temperature=0.2,
        )

        self.chat = self.client.chats.create(
            model="gemini-2.5-flash", config=self.config
        )

    def ask(self, user_input: str) -> str:
        max_retries = 3
        delay = 2

        for attempt in range(max_retries):
            try:
                response = self.chat.send_message(user_input)
                return response.text
            except Exception as e:
                error_msg = str(e)
                if (
                    "503" in error_msg or "UNAVAILABLE" in error_msg
                ) and attempt < max_retries - 1:
                    time.sleep(delay)
                    delay *= 2
                    continue
                return f"Surcotec Agent Error: {error_msg}"