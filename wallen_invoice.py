"""
WALLEN AIR CONDITIONING
Invoice Manager

Version: 1.0 Test Build

Workflow

1. User enters customer street address.
2. Program copies the PDF template.
3. Program creates filename:
       ADDRESS - MMDDYYYY.pdf
4. Program opens PDF.
5. User completes invoice.
6. User saves and closes PDF.
7. Program reads CUST_Address_Street.
8. If address changed, PDF filename is updated.
"""

import logging
import socket
import os
import json
import re
import smtplib
import shutil
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request

from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Callable

from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

from PIL import Image
from PIL import ImageTk

from pypdf import PdfReader


# ============================================================
# APPLICATION PATHS
# ============================================================

SCRIPT_FOLDER = Path(__file__).resolve().parent

COMPANY_FOLDER = SCRIPT_FOLDER

TEMPLATE_FOLDER = COMPANY_FOLDER / "Invoice Template"

TEMPLATE_FILE = (
    TEMPLATE_FOLDER
    / "INVOICE TEMPLATE FILLABLE.pdf"
)

ASSETS_FOLDER = COMPANY_FOLDER / "Assets"

LOGO_FILE = (
    ASSETS_FOLDER
    / "WallenACLogo.png"
)

INVOICE_FOLDER = COMPANY_FOLDER / "Invoices"

# Notes are kept with invoices for now to keep one operational queue.
NOTES_FOLDER = INVOICE_FOLDER

LOG_FOLDER = COMPANY_FOLDER / "Logs"

LOG_FILE = LOG_FOLDER / "WallenInvoice.log"

PRE_OPEN_INVOICE_STATS: dict[str, tuple[int, float]] = {}


# ============================================================
# EMAIL NOTIFICATION CONFIGURATION
# ============================================================

ADMIN_NOTIFICATION_EMAIL = "admin@wallenac.com"

SMTP_SETTINGS_FILE = COMPANY_FOLDER / "smtp_settings.json"

LOCAL_SETTINGS_FOLDER = Path(
    os.getenv("APPDATA", str(COMPANY_FOLDER))
) / "WallenInvoice"

LOCAL_SMTP_SETTINGS_FILE = (
    LOCAL_SETTINGS_FOLDER
    / "smtp_settings.local.json"
)


def read_smtp_settings_file(path: Path) -> dict:

    if not path.exists():
        return {}

    try:
        with open(path, "r", encoding="utf-8") as file:
            data = json.load(file)
    except Exception as ex:
        logging.warning(
            f"Unable to read {path.name}: {ex}"
        )
        return {}

    if not isinstance(data, dict):
        logging.warning(
            f"{path.name} must contain a JSON object"
        )
        return {}

    return data


SHARED_SMTP_SETTINGS = read_smtp_settings_file(
    SMTP_SETTINGS_FILE
)

LOCAL_SMTP_SETTINGS = read_smtp_settings_file(
    LOCAL_SMTP_SETTINGS_FILE
)


def get_smtp_setting(
    json_key: str,
    env_key: str,
    default: str = ""
) -> str:

    env_value = os.getenv(env_key, "").strip()

    if env_value:
        return env_value

    local_value = str(
        LOCAL_SMTP_SETTINGS.get(json_key, "")
    ).strip()

    if local_value:
        return local_value

    file_value = str(
        SHARED_SMTP_SETTINGS.get(json_key, default)
    ).strip()

    return file_value


def parse_smtp_port(raw_port: str) -> int:

    raw_port = raw_port.strip()

    try:
        port = int(raw_port)
    except ValueError:
        logging.warning(
            f"Invalid WALLEN_SMTP_PORT value '{raw_port}'. Using 587."
        )
        return 587

    if port < 1 or port > 65535:
        logging.warning(
            f"Out-of-range WALLEN_SMTP_PORT value '{raw_port}'. Using 587."
        )
        return 587

    return port


EMAIL_SMTP_HOST = get_smtp_setting(
    "smtp_host",
    "WALLEN_SMTP_HOST",
    ""
)

EMAIL_SMTP_PORT = parse_smtp_port(
    get_smtp_setting(
        "smtp_port",
        "WALLEN_SMTP_PORT",
        "587"
    )
)

EMAIL_USE_TLS = (
    get_smtp_setting(
        "smtp_use_tls",
        "WALLEN_SMTP_USE_TLS",
        "true"
    )
    .strip()
    .lower()
    in {"1", "true", "yes", "on"}
)

EMAIL_FROM_ADDRESS = get_smtp_setting(
    "email_from",
    "WALLEN_EMAIL_FROM",
    "noreply@wallenac.com"
).strip()

EMAIL_USERNAME = get_smtp_setting(
    "email_username",
    "WALLEN_EMAIL_USERNAME",
    ""
).strip()

EMAIL_PASSWORD = get_smtp_setting(
    "email_password",
    "WALLEN_EMAIL_PASSWORD",
    ""
)

EMAIL_PROVIDER = get_smtp_setting(
    "email_provider",
    "WALLEN_EMAIL_PROVIDER",
    "auto"
).lower()

GRAPH_TENANT_ID = get_smtp_setting(
    "graph_tenant_id",
    "WALLEN_GRAPH_TENANT_ID",
    ""
)

GRAPH_CLIENT_ID = get_smtp_setting(
    "graph_client_id",
    "WALLEN_GRAPH_CLIENT_ID",
    ""
)

GRAPH_CLIENT_SECRET = get_smtp_setting(
    "graph_client_secret",
    "WALLEN_GRAPH_CLIENT_SECRET",
    ""
)

GRAPH_SCOPE = get_smtp_setting(
    "graph_scope",
    "WALLEN_GRAPH_SCOPE",
    "https://graph.microsoft.com/.default"
)

LAST_SMTP_DIAGNOSTICS = "No SMTP diagnostics captured yet."


# ============================================================
# PDF FIELD DEFINITIONS
# ============================================================

FIELD_CUSTOMER_ADDRESS = "CUST_Address_Street"


# ============================================================
# LOGGING
# ============================================================

def initialize_logging() -> None:

    LOG_FOLDER.mkdir(
        parents=True,
        exist_ok=True
    )

    logging.basicConfig(
        filename=str(LOG_FILE),
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

    logging.info("Application Started")

    if EMAIL_SMTP_HOST:
        logging.info(
            f"SMTP configured for host {EMAIL_SMTP_HOST}:{EMAIL_SMTP_PORT}"
        )
    else:
        logging.info("SMTP host not configured")

    logging.info(
        f"Email provider mode: {EMAIL_PROVIDER}"
    )

    graph_configured = all([
        GRAPH_TENANT_ID,
        GRAPH_CLIENT_ID,
        GRAPH_CLIENT_SECRET,
        EMAIL_FROM_ADDRESS
    ])

    logging.info(
        f"Graph configured: {graph_configured}"
    )

    logging.info(
        f"Shared SMTP settings file: {SMTP_SETTINGS_FILE}"
    )

    logging.info(
        f"Local SMTP settings file: {LOCAL_SMTP_SETTINGS_FILE}"
    )


# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def sanitize_filename(text: str) -> str:
    """
    Removes characters that are not valid
    in Windows file names.
    """

    text = text.strip()

    text = re.sub(
        r'[<>:"/\\|?*]',
        '',
        text
    )

    text = text.replace('#', '')

    text = re.sub(
        r'\s+',
        ' ',
        text
    )

    return text


def current_date_string() -> str:
    """
    Returns MMDDYYYY
    """

    return datetime.now().strftime("%m%d%Y")


# ============================================================
# STARTUP VALIDATION
# ============================================================

def verify_environment() -> bool:

    try:

        INVOICE_FOLDER.mkdir(
            parents=True,
            exist_ok=True
        )

        NOTES_FOLDER.mkdir(
            parents=True,
            exist_ok=True
        )

        LOG_FOLDER.mkdir(
            parents=True,
            exist_ok=True
        )

    except Exception as ex:

        messagebox.showerror(
            "Folder Error",
            f"""
Unable to create required folders.

Error:

{ex}
"""
        )

        return False

    if not TEMPLATE_FOLDER.exists():

        messagebox.showerror(
            "Template Folder Missing",
            f"""
The Invoice Template folder cannot be found.

Expected Location:

{TEMPLATE_FOLDER}

Please contact the office.
"""
        )

        return False

    if not TEMPLATE_FILE.exists():

        messagebox.showerror(
            "Template PDF Missing",
            f"""
The invoice template cannot be found.

Expected File:

{TEMPLATE_FILE}

Please contact the office.
"""
        )

        return False

    logging.info("Environment validation successful")

    return True


# ============================================================
# FILE MANAGEMENT
# ============================================================

def create_unique_file(
    address: str,
    date_text: str | None = None,
    target_folder: Path = INVOICE_FOLDER,
    extension: str = ".pdf"
) -> Path:

    if not date_text:
        date_text = current_date_string()

    if not extension.startswith("."):
        extension = f".{extension}"

    base_name = f"{address} - {date_text}"

    candidate = (
        target_folder
        / f"{base_name}{extension}"
    )

    counter = 2

    while candidate.exists():

        candidate = (
            target_folder
            / f"{base_name} ({counter}){extension}"
        )

        counter += 1

    return candidate


def extract_date_from_filename(pdf_file: Path) -> str | None:

    match = re.search(
        r" - (\d{8})(?: \(\d+\))?\.pdf$",
        pdf_file.name,
        flags=re.IGNORECASE
    )

    if not match:
        return None

    return match.group(1)


def get_file_stats(pdf_file: Path) -> tuple[int, float]:
    try:
        stat = pdf_file.stat()
        return stat.st_size, stat.st_mtime
    except FileNotFoundError:
        return 0, 0.0
    except Exception as ex:
        logging.error(
            f"Unable to read file stats for {pdf_file}: {ex}"
        )
        return 0, 0.0


def capture_invoice_pre_open_stats(pdf_file: Path) -> None:
    size, mtime = get_file_stats(pdf_file)
    PRE_OPEN_INVOICE_STATS[str(pdf_file)] = (size, mtime)
    logging.info(
        f"Pre-open file stats for {pdf_file.name}: size={size} bytes, "
        f"mtime={datetime.fromtimestamp(mtime).isoformat(sep=' ')}"
    )


def is_pdf_file_released(pdf_file: Path) -> bool:
    try:
        with open(pdf_file, "r+b"):
            pass
        return True
    except PermissionError as ex:
        logging.info(
            f"INFO: PDF still open in editor ({pdf_file.name}): {ex}"
        )
        return False
    except OSError as ex:
        if ex.errno in {13, 32, 21, 95}:
            logging.info(
                f"INFO: PDF still open in editor ({pdf_file.name}): {ex}"
            )
            return False
        logging.error(
            f"Unexpected OS condition checking PDF release for {pdf_file.name}: [{ex.errno}] {ex}"
        )
        return False
    except Exception as ex:
        logging.error(
            f"Unexpected exception checking PDF release for {pdf_file.name}: {ex}"
        )
        return False


def wait_for_pdf_release(
    pdf_file: Path,
    on_released: Callable[[], None],
    timeout_seconds: float = 120.0,
    check_interval_ms: int = 1000
) -> None:
    deadline = time.time() + timeout_seconds

    def poll() -> None:
        if not pdf_file.exists():
            set_status("✗ Invoice file missing during release wait")
            logging.warning(
                f"PDF release wait failed, file missing: {pdf_file}"
            )
            enable_ui()
            return

        if is_pdf_file_released(pdf_file):
            logging.info(
                f"PDF release detected for {pdf_file.name}"
            )
            set_status("Invoice closed. Validating saved form fields")
            on_released()
            return

        if time.time() >= deadline:
            set_status("✗ Invoice file still locked after timeout")
            logging.warning(
                f"PDF release timeout for {pdf_file.name}"
            )
            enable_ui()
            return

        root.after(check_interval_ms, poll)

    root.after(check_interval_ms, poll)


def wait_for_file_stability(
    pdf_file: Path,
    consecutive_checks: int = 3,
    interval_seconds: float = 1.0,
    timeout_seconds: float = 60.0
) -> tuple[bool, list[str]]:
    details: list[str] = []
    stable_count = 0
    last_snapshot: tuple[int, float] | None = None
    deadline = time.time() + timeout_seconds

    details.append("SYNC: Starting OneDrive sync monitoring")

    while time.time() < deadline:

        if not pdf_file.exists():
            details.append(
                "FAIL: Invoice file missing during sync monitoring"
            )
            return False, details

        size, mtime = get_file_stats(pdf_file)
        details.append(
            f"SYNC: observed size={size} mtime={datetime.fromtimestamp(mtime).isoformat(sep=' ')}"
        )

        current_snapshot = (size, mtime)

        if last_snapshot is not None and current_snapshot == last_snapshot:
            stable_count += 1
            details.append(
                f"SYNC: unchanged snapshot {stable_count}/{consecutive_checks}"
            )
            if stable_count >= consecutive_checks:
                details.append(
                    "PASS: File metadata remained stable for OneDrive sync"
                )
                return True, details
        else:
            stable_count = 0

        last_snapshot = current_snapshot
        time.sleep(interval_seconds)

    details.append(
        "FAIL: OneDrive synchronization did not stabilize within timeout"
    )
    return False, details


def validate_invoice_integrity(pdf_path: Path) -> tuple[bool, list[str]]:
    report: list[str] = []
    set_status("Validating saved form fields")

    report.append(f"VALIDATION REPORT for {pdf_path}")

    if not pdf_path.exists():
        report.append("FAIL: Invoice file does not exist")
        logging.error("Invoice integrity validation failed: file does not exist")
        logging.info("\n".join(report))
        return False, report

    original_stats = PRE_OPEN_INVOICE_STATS.pop(str(pdf_path), None)
    current_size, current_mtime = get_file_stats(pdf_path)

    if original_stats is None:
        report.append(
            "FAIL: Missing pre-open file metadata needed for integrity validation"
        )
        logging.error(
            "Invoice integrity validation failed: missing pre-open metadata"
        )
        logging.info("\n".join(report))
        return False, report

    original_size, original_mtime = original_stats
    report.append(
        f"PRE-OPEN size={original_size} bytes mtime={datetime.fromtimestamp(original_mtime).isoformat(sep=' ')}"
    )
    report.append(
        f"POST-CLOSE size={current_size} bytes mtime={datetime.fromtimestamp(current_mtime).isoformat(sep=' ')}"
    )

    if current_size == original_size and current_mtime == original_mtime:
        address_value = read_pdf_field_with_retry(
            pdf_path,
            FIELD_CUSTOMER_ADDRESS
        )

        if not address_value:
            report.append(
                f"FAIL: Required PDF field {FIELD_CUSTOMER_ADDRESS} is blank"
            )
            logging.warning(
                "Invoice integrity validation failed: required address field blank"
            )
            logging.info("\n".join(report))
            return False, report

        report.append("PASS: No changes detected after opening the invoice")
        report.append("PASS: Required PDF field {FIELD_CUSTOMER_ADDRESS} contains a value")
        report.append("PASS: Invoice integrity validation completed successfully")
        logging.info("\n".join(report))
        set_status("No changes were made to the invoice.")
        return True, report

    if current_size != original_size:
        report.append("PASS: File size changed after the invoice was closed")
    else:
        report.append("INFO: File size did not change after the invoice was closed")

    if current_mtime != original_mtime:
        report.append(
            "PASS: Modification timestamp changed after the invoice was closed"
        )
    else:
        report.append(
            "INFO: Modification timestamp did not change after the invoice was closed"
        )

    set_status("Checking file integrity")
    set_status("Waiting for OneDrive synchronization")

    stable, sync_details = wait_for_file_stability(pdf_path)
    report.extend(sync_details)

    if not stable:
        report.append("ERROR: OneDrive synchronization failed")
        logging.warning("Invoice integrity validation failed during sync monitoring")
        logging.info("\n".join(report))
        return False, report

    if current_size == 0:
        report.append("FAIL: Invoice file is empty")
        logging.warning("Invoice integrity validation failed: invoice file size is zero")
        logging.info("\n".join(report))
        return False, report

    set_status("Validating saved form fields")
    address_value = read_pdf_field_with_retry(
        pdf_path,
        FIELD_CUSTOMER_ADDRESS
    )

    if not address_value:
        report.append(
            f"FAIL: Required PDF field {FIELD_CUSTOMER_ADDRESS} is blank"
        )
        logging.warning(
            "Invoice integrity validation failed: required address field blank"
        )
        logging.info("\n".join(report))
        return False, report

    report.append(
        f"PASS: Required PDF field {FIELD_CUSTOMER_ADDRESS} contains a value"
    )
    report.append("PASS: Invoice integrity validation completed successfully")
    logging.info("\n".join(report))
    set_status("Verification successful")
    return True, report


# ============================================================
# PDF FUNCTIONS
# ============================================================

def read_pdf_field(
    pdf_file: Path,
    field_name: str
) -> str:

    try:

        reader = PdfReader(
            str(pdf_file)
        )

        fields = reader.get_fields()

        if not fields:
            return ""

        field = fields.get(field_name)

        if not field:
            return ""

        value = field.value

        if value is None:
            return ""

        return str(value).strip()

    except Exception as ex:

        logging.error(
            f"Unable to read field {field_name}: {ex}"
        )

        return ""


def open_pdf(pdf_path: Path) -> None:

    logging.info(
        f"Opening PDF: {pdf_path}"
    )

    os.startfile(str(pdf_path))


def open_text_file(text_path: Path) -> None:

    logging.info(
        f"Opening text file: {text_path}"
    )

    os.startfile(str(text_path))


def read_pdf_field_with_retry(
    pdf_path: Path,
    field_name: str,
    attempts: int = 5,
    delay_seconds: float = 1.0
) -> str:

    for attempt in range(1, attempts + 1):

        value = read_pdf_field(
            pdf_path,
            field_name
        )

        if value:
            return value

        logging.info(
            f"Field {field_name} was blank on attempt {attempt}/{attempts}"
        )

        if attempt < attempts:
            time.sleep(delay_seconds)

    return ""


def send_new_invoice_notification(
    invoice_file: Path,
    entered_address: str,
    final_address: str
) -> bool:
    subject = f"New Invoice Created: {invoice_file.name}"

    body = f"""
A new invoice was created.

File Name: {invoice_file.name}
File Path: {invoice_file}
Created At: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Entered Address: {entered_address}
Final Address: {final_address or '(blank)'}
""".strip()

    success, details = send_email_message(
        subject=subject,
        body=body,
        recipient=ADMIN_NOTIFICATION_EMAIL
    )

    if success:
        logging.info(
            f"Notification email sent for invoice: {invoice_file.name}"
        )
        return True

    if details == "Email notifications not configured":
        logging.info(
            f"Email notifications not configured for {invoice_file.name}."
        )
        return False

    logging.error(
        f"Notification email failed for {invoice_file.name}: {details}"
    )
    return False


def is_graph_configured() -> bool:

    return all([
        GRAPH_TENANT_ID,
        GRAPH_CLIENT_ID,
        GRAPH_CLIENT_SECRET,
        EMAIL_FROM_ADDRESS
    ])


def resolve_email_provider() -> str:

    if EMAIL_PROVIDER == "graph":
        return "graph"

    if EMAIL_PROVIDER == "smtp":
        return "smtp"

    if EMAIL_PROVIDER == "auto":
        if is_graph_configured():
            return "graph"

        if EMAIL_SMTP_HOST:
            return "smtp"

        return "none"

    return "none"


def send_email_via_smtp(
    subject: str,
    body: str,
    recipient: str
) -> tuple[bool, str]:

    if not EMAIL_SMTP_HOST:
        return False, "SMTP host is not configured"

    try:
        message = EmailMessage()
        message["Subject"] = subject
        message["From"] = EMAIL_FROM_ADDRESS
        message["To"] = recipient
        message.set_content(body)

        with smtplib.SMTP(
            EMAIL_SMTP_HOST,
            EMAIL_SMTP_PORT,
            timeout=20
        ) as smtp:

            smtp.ehlo()

            if EMAIL_USE_TLS:
                smtp.starttls()
                smtp.ehlo()

            if EMAIL_USERNAME:
                smtp.login(
                    EMAIL_USERNAME,
                    EMAIL_PASSWORD
                )

            smtp.send_message(message)

        return True, "SMTP send accepted"

    except Exception as ex:
        return False, f"SMTP send failed ({ex})"


def get_graph_access_token() -> tuple[str, str]:

    token_url = (
        f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}/oauth2/v2.0/token"
    )

    payload = urllib.parse.urlencode({
        "client_id": GRAPH_CLIENT_ID,
        "client_secret": GRAPH_CLIENT_SECRET,
        "scope": GRAPH_SCOPE,
        "grant_type": "client_credentials"
    }).encode("utf-8")

    request = urllib.request.Request(
        token_url,
        data=payload,
        headers={
            "Content-Type": "application/x-www-form-urlencoded"
        },
        method="POST"
    )

    try:
        with urllib.request.urlopen(
            request,
            timeout=20
        ) as response:
            token_data = json.loads(
                response.read().decode("utf-8")
            )

        access_token = str(
            token_data.get("access_token", "")
        ).strip()

        if not access_token:
            return "", "Graph token response missing access_token"

        return access_token, ""

    except urllib.error.HTTPError as ex:
        error_text = ex.read().decode("utf-8", errors="replace")
        return "", f"Graph token HTTP {ex.code}: {error_text}"
    except Exception as ex:
        return "", f"Graph token request failed ({ex})"


def send_email_via_graph(
    subject: str,
    body: str,
    recipient: str
) -> tuple[bool, str]:

    if not is_graph_configured():
        return False, "Graph is not fully configured"

    access_token, token_error = get_graph_access_token()

    if not access_token:
        return False, token_error

    endpoint = (
        "https://graph.microsoft.com/v1.0/users/"
        f"{urllib.parse.quote(EMAIL_FROM_ADDRESS)}/sendMail"
    )

    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": recipient
                    }
                }
            ]
        },
        "saveToSentItems": False
    }

    request = urllib.request.Request(
        endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        },
        method="POST"
    )

    try:
        with urllib.request.urlopen(
            request,
            timeout=20
        ):
            pass

        return True, "Graph send accepted"

    except urllib.error.HTTPError as ex:
        error_text = ex.read().decode("utf-8", errors="replace")
        return False, f"Graph send HTTP {ex.code}: {error_text}"
    except Exception as ex:
        return False, f"Graph send failed ({ex})"


def send_email_message(
    subject: str,
    body: str,
    recipient: str
) -> tuple[bool, str]:

    provider = resolve_email_provider()

    if provider == "graph":
        return send_email_via_graph(
            subject,
            body,
            recipient
        )

    if provider == "smtp":
        return send_email_via_smtp(
            subject,
            body,
            recipient
        )

    return False, "Email notifications not configured"


def set_last_smtp_diagnostics(steps: list[str]) -> None:

    global LAST_SMTP_DIAGNOSTICS

    if not steps:
        LAST_SMTP_DIAGNOSTICS = "No SMTP diagnostics captured yet."
        return

    LAST_SMTP_DIAGNOSTICS = "\n".join(steps)


def set_smtp_health(
    text: str,
    color: str
) -> None:

    try:
        smtp_health_var.set(text)
        smtp_health_dot_label.config(fg=color)
        smtp_health_value_label.config(fg=color)
    except NameError:
        pass


def run_smtp_connectivity_check() -> tuple[bool, list[str]]:

    steps: list[str] = []

    if not EMAIL_SMTP_HOST:
        steps.append("INFO: WALLEN_SMTP_HOST is not configured")
        return False, steps

    try:
        socket.getaddrinfo(
            EMAIL_SMTP_HOST,
            EMAIL_SMTP_PORT
        )
        steps.append(f"PASS: DNS resolved {EMAIL_SMTP_HOST}")
    except Exception as ex:
        steps.append(f"FAIL: DNS resolution failed ({ex})")
        return False, steps

    try:
        with smtplib.SMTP(
            EMAIL_SMTP_HOST,
            EMAIL_SMTP_PORT,
            timeout=20
        ) as smtp:
            steps.append(
                f"PASS: Connected to {EMAIL_SMTP_HOST}:{EMAIL_SMTP_PORT}"
            )

            smtp.ehlo()
            steps.append("PASS: EHLO completed")

            if EMAIL_USE_TLS:
                smtp.starttls()
                smtp.ehlo()
                steps.append("PASS: STARTTLS negotiated")
            else:
                steps.append("PASS: TLS disabled by configuration")

            if EMAIL_USERNAME:
                smtp.login(
                    EMAIL_USERNAME,
                    EMAIL_PASSWORD
                )
                steps.append("PASS: SMTP authentication succeeded")
            else:
                steps.append("PASS: Authentication skipped (no username)")

        return True, steps

    except Exception as ex:
        steps.append(f"FAIL: Connectivity check failed ({ex})")
        return False, steps


def run_smtp_self_test() -> tuple[bool, list[str]]:

    success, steps = run_smtp_connectivity_check()

    if not success:
        return False, steps

    test_id = datetime.now().strftime("%Y%m%d-%H%M%S")

    try:
        message = EmailMessage()
        message["Subject"] = f"SMTP Test {test_id} - Wallen Invoice Manager"
        message["From"] = EMAIL_FROM_ADDRESS
        message["To"] = ADMIN_NOTIFICATION_EMAIL
        message.set_content(
            f"""
SMTP self-test succeeded.

Test ID: {test_id}
Host: {EMAIL_SMTP_HOST}
Port: {EMAIL_SMTP_PORT}
TLS Enabled: {EMAIL_USE_TLS}
From: {EMAIL_FROM_ADDRESS}
Machine: {os.getenv('COMPUTERNAME', 'Unknown')}
""".strip()
        )

        with smtplib.SMTP(
            EMAIL_SMTP_HOST,
            EMAIL_SMTP_PORT,
            timeout=20
        ) as smtp:

            smtp.ehlo()

            if EMAIL_USE_TLS:
                smtp.starttls()
                smtp.ehlo()

            if EMAIL_USERNAME:
                smtp.login(
                    EMAIL_USERNAME,
                    EMAIL_PASSWORD
                )

            smtp.send_message(message)

        steps.append(
            f"PASS: Test email accepted for delivery to {ADMIN_NOTIFICATION_EMAIL}"
        )

        return True, steps

    except Exception as ex:
        steps.append(f"FAIL: SMTP test failed ({ex})")
        return False, steps


def run_graph_connectivity_check() -> tuple[bool, list[str]]:

    steps: list[str] = []

    if not is_graph_configured():
        steps.append("INFO: Graph provider is not fully configured")
        return False, steps

    access_token, token_error = get_graph_access_token()

    if not access_token:
        steps.append(f"FAIL: {token_error}")
        return False, steps

    steps.append("PASS: Graph token acquired")
    return True, steps


def run_email_connectivity_check() -> tuple[bool, list[str]]:

    provider = resolve_email_provider()

    if provider == "graph":
        success, steps = run_graph_connectivity_check()
        return success, ["INFO: Provider graph", *steps]

    if provider == "smtp":
        success, steps = run_smtp_connectivity_check()
        return success, ["INFO: Provider smtp", *steps]

    return False, ["INFO: No email provider configured"]


def run_email_self_test() -> tuple[bool, list[str]]:

    success, steps = run_email_connectivity_check()

    if not success:
        return False, steps

    provider = resolve_email_provider()
    test_id = datetime.now().strftime("%Y%m%d-%H%M%S")

    subject = f"Email Test {test_id} - Wallen Invoice Manager"
    body = f"""
Email self-test succeeded.

Test ID: {test_id}
Provider: {provider}
From: {EMAIL_FROM_ADDRESS}
Machine: {os.getenv('COMPUTERNAME', 'Unknown')}
""".strip()

    sent, details = send_email_message(
        subject=subject,
        body=body,
        recipient=ADMIN_NOTIFICATION_EMAIL
    )

    if sent:
        steps.append(
            f"PASS: Test email accepted for delivery to {ADMIN_NOTIFICATION_EMAIL}"
        )
        return True, steps

    steps.append(f"FAIL: {details}")
    return False, steps


def run_startup_smtp_health_check() -> None:

    success, steps = run_email_connectivity_check()
    set_last_smtp_diagnostics(steps)

    details = "\n".join(steps)

    if success:
        logging.info(
            f"Startup email connectivity check passed. Steps: {details}"
        )
        root.after(
            0,
            lambda: set_smtp_health("Healthy", "#1B7F3B")
        )
        return

    if any(step.startswith("INFO:") for step in steps):
        logging.info(
            f"Startup email connectivity check skipped. Steps: {details}"
        )
        root.after(
            0,
            lambda: set_smtp_health("Not configured", "#667085")
        )
        return

    logging.warning(
        f"Startup email connectivity check failed. Steps: {details}"
    )
    root.after(
        0,
        lambda: set_smtp_health("Issue detected", "#B42318")
    )


def start_startup_smtp_health_check() -> None:

    set_smtp_health("Checking...", "#B54708")

    worker = threading.Thread(
        target=run_startup_smtp_health_check,
        daemon=True
    )
    worker.start()


def copy_smtp_diagnostics() -> None:

    diagnostics_text = LAST_SMTP_DIAGNOSTICS

    root.clipboard_clear()
    root.clipboard_append(diagnostics_text)
    root.update()

    set_status("✓ Email diagnostics copied")


def test_email_settings() -> None:

    set_status("Running email test...")
    disable_ui()
    root.update_idletasks()

    success, steps = run_email_self_test()
    set_last_smtp_diagnostics(steps)

    enable_ui()

    details = "\n".join(steps)

    if success:

        logging.info(
            f"Email self-test passed. Steps: {details}"
        )

        set_status("✓ Email test passed")
        set_smtp_health("Healthy", "#1B7F3B")

        messagebox.showinfo(
            "Email Test Passed",
            f"""
Email self-test completed successfully.

{details}

Please confirm receipt in {ADMIN_NOTIFICATION_EMAIL}.
"""
        )

        return

    logging.error(
        f"Email self-test failed. Steps: {details}"
    )

    set_status("✗ Email test failed")
    set_smtp_health("Issue detected", "#B42318")

    messagebox.showerror(
        "Email Test Failed",
        f"""
Email self-test did not complete.

{details}
"""
    )


# ============================================================
# INVOICE CREATION
# ============================================================

def find_matching_invoices(address: str) -> list[Path]:

    normalized_address = sanitize_filename(address).lower()

    if not normalized_address:
        return []

    matches: list[Path] = []

    for pdf_file in INVOICE_FOLDER.glob("*.pdf"):

        invoice_name = sanitize_filename(pdf_file.stem).lower()

        if invoice_name.startswith(normalized_address):
            matches.append(pdf_file)

    return matches


def select_most_recent_invoice(invoices: list[Path]) -> Path | None:

    if not invoices:
        return None

    return max(
        invoices,
        key=lambda invoice: invoice.stat().st_mtime
    )


def handle_address_enter(event=None) -> None:

    address = address_var.get().strip()

    if not address:
        set_status(
            "⚠ Enter customer street address (e.g., 1234 SW Palm Beach Drive)"
        )
        address_entry.focus_set()
        return

    matching_invoices = find_matching_invoices(address)

    if not matching_invoices:
        set_status("No existing invoice found. Creating new invoice...")
        create_new_invoice()
        return

    invoice_to_open = select_most_recent_invoice(matching_invoices)

    if invoice_to_open is None:
        set_status("No existing invoice found. Creating new invoice...")
        create_new_invoice()
        return

    set_status("Existing invoice found. Opening...")
    edit_existing_invoice(invoice_to_open)


def create_new_invoice() -> None:

    address = (
        address_var.get()
        .strip()
    )

    if not address:

        set_status("⚠ Enter customer street address (e.g., 1234 SW Palm Beach Drive)")
        address_entry.focus_set()

        return

    address = sanitize_filename(address)

    pdf_file = create_unique_file(address)

    try:

        shutil.copy2(
            TEMPLATE_FILE,
            pdf_file
        )

    except Exception as ex:

        set_status(f"✗ Copy Error: {ex}")
        logging.error(
            f"Copy Failed: {ex}"
        )

        return

    logging.info(
        f"New Invoice Created: {pdf_file.name}"
    )

    set_status("Invoice opened. Waiting for PDF to close...")
    capture_invoice_pre_open_stats(pdf_file)
    disable_ui()
    open_pdf(pdf_file)

    wait_for_pdf_release(
        pdf_file,
        lambda: finish_create_invoice(pdf_file, address)
    )


def finish_create_invoice(pdf_file: Path, address: str) -> None:
    final_address = read_pdf_field_with_retry(
        pdf_file,
        FIELD_CUSTOMER_ADDRESS
    )

    if not final_address:
        original_stats = PRE_OPEN_INVOICE_STATS.get(str(pdf_file))

        if original_stats is not None:
            current_size, current_mtime = get_file_stats(pdf_file)
            original_size, original_mtime = original_stats

            if current_size == original_size and current_mtime == original_mtime:
                keep_blank = messagebox.askyesno(
                    "Blank Invoice Detected",
                    "No invoice data was detected.\n\n"
                    "The invoice appears to have been closed without entering information.\n\n"
                    "Delete the blank invoice?"
                )

                if keep_blank:
                    try:
                        pdf_file.unlink(missing_ok=True)
                        logging.info(
                            "Blank invoice removed by user confirmation"
                        )
                        set_status("Invoice cancelled.")
                    except Exception as ex:
                        logging.error(
                            f"Blank invoice delete failed: {ex}"
                        )
                        set_status(f"✗ Delete failed: {ex}")
                        enable_ui()
                        return
                else:
                    logging.info(
                        "Blank invoice retained by user choice"
                    )
                    set_status("Blank invoice retained.")

                refresh_invoice_count()
                address_var.set("")
                address_entry.focus_set()
                enable_ui()
                return

        logging.warning(
            f"Final address field was blank for {pdf_file.name}"
        )

    integrity_success, integrity_report = validate_invoice_integrity(
        pdf_file
    )

    if not integrity_success:
        logging.warning(
            f"Invoice integrity validation failed for {pdf_file.name}: "
            f"{integrity_report[-1] if integrity_report else 'unknown reason'}"
        )
        set_status(
            "✗ Invoice validation failed. Reopen the invoice and save again."
        )
        enable_ui()
        return

    if final_address:
        final_address = sanitize_filename(final_address)

        if final_address != address:
            try:
                replacement = create_unique_file(
                    final_address
                )

                pdf_file.rename(replacement)
                logging.info(
                    f"Renamed: {pdf_file.name} -> {replacement.name}"
                )
                pdf_file = replacement
                set_status(f"✓ Address updated: {pdf_file.name}")
            except Exception as ex:
                logging.error(
                    f"Rename Failed: {ex}"
                )
                set_status(f"✗ Rename Error: {ex}")
                enable_ui()
                return

    email_sent = send_new_invoice_notification(
        invoice_file=pdf_file,
        entered_address=address,
        final_address=final_address
    )

    refresh_invoice_count()
    address_var.set("")
    address_entry.focus_set()

    if email_sent:
        set_status(f"✓ Saved + emailed: {pdf_file.name}")
    else:
        set_status(f"✓ Saved (email skipped/failed): {pdf_file.name}")

    enable_ui()


def finish_edit_invoice(pdf_file: Path, original_address: str) -> None:
    integrity_success, integrity_report = validate_invoice_integrity(
        pdf_file
    )

    if not integrity_success:
        logging.warning(
            f"Invoice integrity validation failed for {pdf_file.name}: "
            f"{integrity_report[-1] if integrity_report else 'unknown reason'}"
        )
        set_status(
            "✗ Invoice validation failed. Reopen the invoice and save again."
        )
        enable_ui()
        return

    current_address = read_pdf_field_with_retry(
        pdf_file,
        FIELD_CUSTOMER_ADDRESS
    )

    if not current_address:
        set_status("✓ Invoice updated")
        address_var.set("")
        address_entry.focus_set()
        enable_ui()
        return

    current_address = sanitize_filename(current_address)
    original_address = sanitize_filename(original_address)

    if current_address == original_address:
        set_status("✓ Invoice updated")
        address_var.set("")
        address_entry.focus_set()
        enable_ui()
        return

    try:
        date_text = extract_date_from_filename(pdf_file)
        replacement = create_unique_file(
            current_address,
            date_text
        )

        pdf_file.rename(replacement)

        logging.info(
            f"Invoice Renamed: {replacement.name}"
        )

        set_status(f"✓ Renamed: {replacement.name}")
    except Exception as ex:
        logging.error(
            f"Rename Error: {ex}"
        )
        set_status(f"✗ Rename Error: {ex}")

    enable_ui()


def create_new_note() -> None:

    address = (
        address_var.get()
        .strip()
    )

    if not address:

        set_status("⚠ Enter customer street address before creating a note")
        address_entry.focus_set()

        return

    address = sanitize_filename(address)

    note_file = create_unique_file(
        address=address,
        target_folder=NOTES_FOLDER,
        extension=".txt"
    )

    try:
        note_file.touch(
            exist_ok=False
        )
    except Exception as ex:
        set_status(f"✗ Note Error: {ex}")
        logging.error(
            f"Note creation failed: {ex}"
        )
        return

    logging.info(
        f"New Note Created: {note_file.name}"
    )

    try:
        open_text_file(note_file)
    except Exception as ex:
        set_status(f"✗ Open Note Error: {ex}")
        logging.error(
            f"Open note failed: {ex}"
        )
        return

    address_var.set("")
    address_entry.focus_set()
    set_status(f"✓ Note ready: {note_file.name}")


# ============================================================
# EDIT EXISTING INVOICE
# ============================================================

def edit_existing_invoice(invoice_path: Path | None = None) -> None:

    if invoice_path is None:
        selected = filedialog.askopenfilename(
            title="Select Existing Invoice",
            initialdir=str(INVOICE_FOLDER),
            filetypes=[
                (
                    "PDF Files",
                    "*.pdf"
                )
            ]
        )

        if not selected:
            return

        pdf_file = Path(selected)
    else:
        pdf_file = invoice_path

        if not pdf_file.exists():
            set_status("✗ Invoice not found")
            return

    logging.info(
        f"Editing Existing Invoice: {pdf_file.name}"
    )

    original_address = read_pdf_field(
        pdf_file,
        FIELD_CUSTOMER_ADDRESS
    )

    set_status("Invoice opened. Waiting for PDF to close...")
    capture_invoice_pre_open_stats(pdf_file)
    disable_ui()
    open_pdf(pdf_file)

    wait_for_pdf_release(
        pdf_file,
        lambda: finish_edit_invoice(pdf_file, original_address)
    )


# ============================================================
# GUI
# ============================================================

def invoice_count() -> int:

    return sum(
        1
        for _ in INVOICE_FOLDER.glob("*.pdf")
    )


def set_status(text: str) -> None:

    try:
        status_var.set(text)
    except NameError:
        pass


def refresh_invoice_count() -> None:

    try:
        invoice_count_label.config(
            text=f"Active Invoices Found: {invoice_count()}"
        )
    except NameError:
        pass


def disable_ui() -> None:

    try:
        create_button.config(state="disabled")
        edit_button.config(state="disabled")
        note_button.config(state="disabled")
        test_email_button.config(state="disabled")
        copy_smtp_button.config(state="disabled")
        address_entry.config(state="disabled")
    except NameError:
        pass


def enable_ui() -> None:

    try:
        create_button.config(state="normal")
        edit_button.config(state="normal")
        note_button.config(state="normal")
        test_email_button.config(state="normal")
        copy_smtp_button.config(state="normal")
        address_entry.config(state="normal")
    except NameError:
        pass


# ============================================================
# APPLICATION START
# ============================================================

root = Tk()

root.title(
    "Wallen Invoice Manager"
)

root.geometry(
    "860x820"
)

root.minsize(
    760,
    740
)

root.resizable(
    True,
    True
)

initialize_logging()

if not verify_environment():

    root.destroy()

    sys.exit()

try:

    logo_image = Image.open(LOGO_FILE)

    logo_image.thumbnail(
        (500, 180),
        Image.LANCZOS
    )

    logo_photo = ImageTk.PhotoImage(
        logo_image
    )

    logo_label = Label(
        root,
        image=logo_photo
    )

    logo_label.pack(
        pady=(10, 5)
    )

except Exception:

    Label(
        root,
        text="WALLEN AIR CONDITIONING",
        font=("Segoe UI", 18, "bold")
    ).pack(
        pady=(15, 5)
    )

Label(
    root,
    text="""
INVOICE MANAGER

Enter the customer street address:

Example:  1234 SW Palm Beach Drive

(STREET ADDRESS ONLY - no City, State, ZIP)
""",
    justify=CENTER
).pack(
    pady=10
)

address_var = StringVar()

address_entry = Entry(
    root,
    width=60,
    textvariable=address_var,
    font=("Segoe UI", 11)
)

address_entry.pack(
    pady=10
)

address_entry.bind("<Return>", handle_address_enter)

invoice_count_label = Label(
    root,
    text=f"Active Invoices Found: {invoice_count()}"
)

invoice_count_label.pack(
    pady=5
)

status_var = StringVar()

status_var.set("Ready")

status_frame = Frame(
    root,
    bd=1,
    relief="sunken"
)

status_frame.pack(
    fill="x",
    padx=20,
    pady=10
)

Label(
    status_frame,
    text="Status:",
    font=("Segoe UI", 10, "bold")
).pack(
    anchor="w"
)

Label(
    status_frame,
    textvariable=status_var
).pack(
    anchor="w"
)

smtp_health_row = Frame(status_frame)
smtp_health_row.pack(
    anchor="w",
    pady=(4, 0)
)

Label(
    smtp_health_row,
    text="Email Health:",
    font=("Segoe UI", 10, "bold")
).pack(
    side="left"
)

smtp_health_dot_label = Label(
    smtp_health_row,
    text=" ● ",
    fg="#B54708"
)
smtp_health_dot_label.pack(
    side="left"
)

smtp_health_var = StringVar()
smtp_health_var.set("Checking...")

smtp_health_value_label = Label(
    smtp_health_row,
    textvariable=smtp_health_var,
    fg="#B54708"
)
smtp_health_value_label.pack(
    side="left"
)

BUTTON_WIDTH = 40
BUTTON_FONT = ("Segoe UI", 11, "bold")
SHOW_EMAIL_TOOLS = False

create_button = Button(
    root,
    text="Create New Invoice",
    width=BUTTON_WIDTH,
    font=BUTTON_FONT,
    command=create_new_invoice
)
create_button.pack(
    pady=5
)

edit_button = Button(
    root,
    text="Edit Existing Invoice",
    width=BUTTON_WIDTH,
    font=BUTTON_FONT,
    command=edit_existing_invoice
)
edit_button.pack(
    pady=5
)

note_button = Button(
    root,
    text="Create New Note",
    width=BUTTON_WIDTH,
    font=BUTTON_FONT,
    command=create_new_note
)
note_button.pack(
    pady=5
)

if SHOW_EMAIL_TOOLS:

    test_email_button = Button(
        root,
        text="Test Email Settings",
        width=BUTTON_WIDTH,
        font=BUTTON_FONT,
        command=test_email_settings
    )
    test_email_button.pack(
        pady=5
    )

    copy_smtp_button = Button(
        root,
        text="Copy Email Diagnostics",
        width=BUTTON_WIDTH,
        font=BUTTON_FONT,
        command=copy_smtp_diagnostics
    )
    copy_smtp_button.pack(
        pady=5
    )

Button(
    root,
    text="Exit",
    width=BUTTON_WIDTH,
    font=BUTTON_FONT,
    command=root.destroy
).pack(
    pady=5
)

root.after(
    400,
    start_startup_smtp_health_check
)

root.mainloop()