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
import os
import re
import shutil
import sys
import time

from datetime import datetime
from pathlib import Path

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

PROCESSED_FOLDER = COMPANY_FOLDER / "Processed"

LOG_FOLDER = COMPANY_FOLDER / "Logs"

LOG_FILE = LOG_FOLDER / "WallenInvoice.log"


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

def create_unique_file(address: str) -> Path:

    date_text = current_date_string()

    base_name = f"{address} - {date_text}"

    candidate = (
        INVOICE_FOLDER
        / f"{base_name}.pdf"
    )

    counter = 2

    while candidate.exists():

        candidate = (
            INVOICE_FOLDER
            / f"{base_name} ({counter}).pdf"
        )

        counter += 1

    return candidate


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


def wait_for_pdf_close(
    pdf_path: Path,
    timeout_hours: int = 8
) -> bool:

    start_time = time.time()

    timeout_seconds = (
        timeout_hours
        * 60
        * 60
    )

    while True:

        if (
            time.time() - start_time
            > timeout_seconds
        ):

            logging.warning(
                f"Timeout waiting for {pdf_path}"
            )

            return False

        try:

            with open(pdf_path, "ab"):
                return True

        except PermissionError:

            time.sleep(2)

        except Exception:

            time.sleep(2)


# ============================================================
# INVOICE CREATION
# ============================================================

def create_new_invoice() -> None:

    address = (
        address_var.get()
        .strip()
    )

    if not address:

        messagebox.showwarning(
            "Address Required",
            """
Please enter the customer street address.

Example:

1234 SW Palm Beach Drive

Do NOT enter:

City
State
ZIP Code
"""
        )

        return

    address = sanitize_filename(address)

    pdf_file = create_unique_file(address)

    try:

        shutil.copy2(
            TEMPLATE_FILE,
            pdf_file
        )

    except Exception as ex:

        messagebox.showerror(
            "Copy Error",
            str(ex)
        )

        logging.error(
            f"Copy Failed: {ex}"
        )

        return

    logging.info(
        f"New Invoice Created: {pdf_file.name}"
    )

    messagebox.showinfo(
        "Step 2 of 3",
        f"""
Customer Street Address:

{address}

Invoice File Name:

{pdf_file.name}

Click OK to open the invoice.
"""
    )

    messagebox.showinfo(
        "Step 3 of 3",
        """
The invoice is now open.

Please do the following:

1. Complete the invoice.
2. Save the invoice.
3. Close the PDF.

The program will continue automatically.
"""
    )

    set_status("Opening invoice...")
    disable_ui()
    hide_root()
    open_pdf(pdf_file)

    closed = wait_for_pdf_close(pdf_file)

    restore_root()
    enable_ui()
    set_status("Invoice window closed")

    if not closed:

        messagebox.showwarning(
            "Timeout",
            """
The invoice remained open too long.

Please reopen the application if needed.
"""
        )

        return

    final_address = read_pdf_field(
        pdf_file,
        FIELD_CUSTOMER_ADDRESS
    )

    if final_address:

        final_address = sanitize_filename(
            final_address
        )

        if final_address != address:

            try:

                replacement = (
                    create_unique_file(
                        final_address
                    )
                )

                pdf_file.rename(
                    replacement
                )

                logging.info(
                    f"Renamed: {pdf_file.name} -> {replacement.name}"
                )

                pdf_file = replacement

                messagebox.showinfo(
                    "Address Updated",
                    f"""
Customer address changed.

New File Name:

{pdf_file.name}
"""
                )

            except Exception as ex:

                logging.error(
                    f"Rename Failed: {ex}"
                )

    messagebox.showinfo(
        "Invoice Complete",
        f"""
Invoice processing complete.

File Location:

{pdf_file}
"""
    )

    refresh_invoice_count()
    address_var.set("")
    address_entry.focus_set()
    set_status(
        f"Last Invoice Saved: {pdf_file.name}"
    )


# ============================================================
# EDIT EXISTING INVOICE
# ============================================================

def edit_existing_invoice() -> None:

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

    logging.info(
        f"Editing Existing Invoice: {pdf_file.name}"
    )

    original_address = read_pdf_field(
        pdf_file,
        FIELD_CUSTOMER_ADDRESS
    )

    messagebox.showinfo(
        "Edit Instructions",
        """
The invoice is now open.

Please:

1. Make your changes.
2. Save the invoice.
3. Close the PDF.

The program will continue automatically.
"""
    )

    set_status("Opening existing invoice...")
    disable_ui()
    hide_root()
    open_pdf(pdf_file)

    closed = wait_for_pdf_close(pdf_file)

    restore_root()
    enable_ui()
    set_status("Invoice window closed")

    if not closed:
        set_status("Invoice wait timed out")
        return

    current_address = read_pdf_field(
        pdf_file,
        FIELD_CUSTOMER_ADDRESS
    )

    if not current_address:
        return

    current_address = sanitize_filename(
        current_address
    )

    original_address = sanitize_filename(
        original_address
    )

    if current_address == original_address:
        return

    try:

        replacement = create_unique_file(
            current_address
        )

        pdf_file.rename(
            replacement
        )

        logging.info(
            f"Invoice Renamed: {replacement.name}"
        )

        messagebox.showinfo(
            "Invoice Renamed",
            f"""
Customer address was updated.

New File Name:

{replacement.name}
"""
        )

    except Exception as ex:

        logging.error(
            f"Rename Error: {ex}"
        )

        messagebox.showerror(
            "Rename Error",
            str(ex)
        )


# ============================================================
# GUI
# ============================================================

def invoice_count() -> int:

    return len(
        list(
            INVOICE_FOLDER.glob("*.pdf")
        )
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
        address_entry.config(state="disabled")
    except NameError:
        pass


def enable_ui() -> None:

    try:
        create_button.config(state="normal")
        edit_button.config(state="normal")
        address_entry.config(state="normal")
    except NameError:
        pass


def hide_root() -> None:

    root.withdraw()
    root.update()


def restore_root() -> None:

    root.deiconify()
    root.lift()
    root.focus_force()


# ============================================================
# APPLICATION START
# ============================================================

root = Tk()

root.title(
    "Wallen Invoice Manager"
)

root.geometry(
    "750x600"
)

root.resizable(
    False,
    False
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
What would you like to do?

To Create a New Invoice:

Enter the customer street address.

Example:

1234 SW Palm Beach Drive

Enter STREET ADDRESS ONLY.
Do NOT enter City, State, or ZIP Code.
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

Button(
    root,
    text="Create New Invoice",
    width=30,
    height=2,
    command=create_new_invoice
).pack(
    pady=5
)

Button(
    root,
    text="Open Existing Invoice",
    width=30,
    height=2,
    command=edit_existing_invoice
).pack(
    pady=5
)

Button(
    root,
    text="Exit",
    width=30,
    height=2,
    command=root.destroy
).pack(
    pady=5
)

root.mainloop()