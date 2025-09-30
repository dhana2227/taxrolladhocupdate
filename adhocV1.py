import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import pyperclip
import pyodbc
from datetime import datetime
import getpass
import ctypes
import threading
import os
import tempfile
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import re
from PIL import Image, ImageTk
import pickle

# Global session updates
session_user_authenticated = False
authenticated_user = None

# File to store authentication data
AUTH_FILE = os.path.join(tempfile.gettempdir(), ".taxroll_auth_cache")

session_updates = {
    "Value Update": [],
    "LUC Update": [],
    "Landsize Update": [],
    "GBA Update": [],
    "Taxroll Insert": []
}

# Headers for Excel report
headers_map = {
    "Value Update": ["Taxyear", "CADID", "AccountNumber", "OCALUC", "LandValue", "ImprovementValue", "MarketValues",
                     "AssessedValues", "Batch", "ValueType"],
    "LUC Update": ["Taxyear", "CADID", "AccountNumber", "PreviousCADLUC", "Update_CADLUC", "Update_Taxroll_LUC",
                   "Update_OCALUC", "OCA_Desc", "Category", "Batch"],
    "Landsize Update": ["Taxyear", "CADID", "AccountNumber", "LandSQFT", "Batch"],
    "GBA Update": ["Taxyear", "CADID", "AccountNumber", "GBA", "NRA", "YearBuilt", "Batch", "Remarks"],
    "Taxroll Insert": ["CADAccountNumber", "Year", "LegalDescription", "ParcelID", "ClassCode", "Remarks",
                       "LandSize", "LandUseCode", "NeighborhoodCode", "ExemptionCode", "GBA", "NRA", "Units", "Grade",
                       "StreetNumber", "StreetName", "NoticedLandValue", "NoticedImprovedValue", "NoticedTotalValue",
                       "YearBuilt",
                       "OwnerName", "OwnerAddress", "OwnerAddress2", "OwnerAddress3", "OwnerCitySt", "OwnerCity",
                       "OwnerState",
                       "OwnerZip", "OwnerZip4", "NoticedMarketValue", "Keymap", "EconomicArea", "SubDivi",
                       "PropAddress", "PropCity",
                       "PropZip", "PropZip4", "CadID", "agentcode", "Agent_Name", "IsUdiAccount", "Ag_Value",
                       "Dump_Landusecode",
                       "Township", "Noticed_Date", "Farmland_Value", "Farmbuilding_value", "Revenue",
                       "Hotel_Classification",
                       "PropertyName"
                       ]
}

# Email and DB config
SMTP_SERVER = "pfbaexch.poconnor.com"
EMAIL_FROM = "Dhanasekaranr@pathfinderanalysis.com"
EMAIL_TO = "Dhanasekaranr@pathfinderanalysis.com"

DB_CONNECTIONS = [
    r'DRIVER={ODBC Driver 13 for SQL Server};SERVER=taxrollstage-db;DATABASE=TaxrollStaging;Trusted_Connection=yes;',
    r'DRIVER={ODBC Driver 13 for SQL Server};SERVER=azspxes;DATABASE=TaxrollStaging;Trusted_Connection=yes;'
]

# Color scheme
COLORS = {
    'primary': '#2E86AB',
    'secondary': '#A23B72',
    'accent': '#F18F01',
    'success': '#4CAF50',
    'warning': '#FF9800',
    'error': '#F44336',
    'light_bg': '#F8F9FA',
    'dark_bg': '#343A40',
    'white': '#FFFFFF',
    'text_dark': '#212529',
    'text_light': '#6C757D'
}


def save_auth_data(username):
    """Save authentication data to file"""
    try:
        auth_data = {
            'username': username,
            'timestamp': datetime.now().isoformat(),
            'machine': os.environ.get('COMPUTERNAME', 'unknown')
        }
        with open(AUTH_FILE, 'wb') as f:
            pickle.dump(auth_data, f)
    except Exception as e:
        print(f"Failed to save auth data: {e}")


def load_auth_data():
    """Load authentication data from file"""
    try:
        if os.path.exists(AUTH_FILE):
            with open(AUTH_FILE, 'rb') as f:
                auth_data = pickle.load(f)
                # Check if auth data is from same machine and recent (within 7 days)
                if auth_data.get('machine') == os.environ.get('COMPUTERNAME', 'unknown'):
                    saved_time = datetime.fromisoformat(auth_data['timestamp'])
                    if (datetime.now() - saved_time).days < 7:
                        return auth_data.get('username')
    except Exception as e:
        print(f"Failed to load auth data: {e}")
    return None


def clear_auth_data():
    """Clear saved authentication data"""
    try:
        if os.path.exists(AUTH_FILE):
            os.remove(AUTH_FILE)
    except Exception as e:
        print(f"Failed to clear auth data: {e}")


def verify_windows_login(username, password):
    """Enhanced Windows authentication"""
    try:
        if not password or not username:
            return False
        LOGON32_LOGON_INTERACTIVE = 2
        LOGON32_PROVIDER_DEFAULT = 0
        advapi32 = ctypes.windll.advapi32
        handle = ctypes.c_void_p()
        result = advapi32.LogonUserW(username, None, password,
                                     LOGON32_LOGON_INTERACTIVE,
                                     LOGON32_PROVIDER_DEFAULT,
                                     ctypes.byref(handle))
        if result != 0:
            ctypes.windll.kernel32.CloseHandle(handle)
            return True
        return False
    except Exception as e:
        print(f"Authentication error: {e}")
        return False


def insert_to_servers(query, values):
    """Insert data to multiple database servers"""
    success_count = 0
    for i, conn_str in enumerate(DB_CONNECTIONS):
        try:
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            cursor.execute(query, values)
            conn.commit()
            cursor.close()
            conn.close()
            success_count += 1
        except Exception as e:
            print(f"Database {i + 1} insert failed: {e}")
    return success_count


def sanitize_row(row, expected_len, numeric_cols=None):
    """Clean and validate row data"""
    if numeric_cols is None:
        numeric_cols = set()

    cleaned = []
    for i, val in enumerate(row):
        val = str(val).strip() if val else None
        if not val:
            cleaned.append(None)
        elif i in numeric_cols:
            val = re.sub(r'[^0-9.-]', '', val)
            try:
                cleaned.append(float(val) if val else None)
            except ValueError:
                cleaned.append(None)
        else:
            cleaned.append(val)

    while len(cleaned) < expected_len:
        cleaned.append(None)
    return cleaned


def export_to_excel_multi_sheet(data_dict, file_path, headers_map):
    """Export data to Excel with multiple sheets"""
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for module_name, rows in data_dict.items():
            if rows:
                headers = headers_map[module_name] + ["CreatedBy", "CreatedDate"]
                df = pd.DataFrame(rows, columns=headers)
                df.to_excel(writer, sheet_name=module_name[:31], index=False)


def send_email_with_attachment(subject, body, attachment_path, html=False):
    """Send email with attachment"""
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_FROM
        msg['To'] = EMAIL_TO
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html' if html else 'plain'))

        with open(attachment_path, "rb") as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(attachment_path)}"')
        msg.attach(part)

        smtp = smtplib.SMTP(SMTP_SERVER)
        smtp.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())
        smtp.quit()
        return True
    except Exception as e:
        print(f"Email sending failed: {e}")
        return False


class ModernButton(tk.Button):
    """Custom modern button with hover effects"""

    def __init__(self, parent, text="", command=None, style="primary", **kwargs):
        # Style configurations
        styles = {
            "primary": {"bg": COLORS['primary'], "fg": COLORS['white'], "hover": "#1E5F7A"},
            "success": {"bg": COLORS['success'], "fg": COLORS['white'], "hover": "#45A049"},
            "warning": {"bg": COLORS['warning'], "fg": COLORS['white'], "hover": "#E68900"},
            "error": {"bg": COLORS['error'], "fg": COLORS['white'], "hover": "#D32F2F"},
            "secondary": {"bg": COLORS['light_bg'], "fg": COLORS['text_dark'], "hover": "#E2E6EA"}
        }

        style_config = styles.get(style, styles["primary"])

        super().__init__(
            parent,
            text=text,
            command=command,
            bg=style_config["bg"],
            fg=style_config["fg"],
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            bd=0,
            padx=20,
            pady=10,
            cursor="hand2",
            **kwargs
        )

        self.default_bg = style_config["bg"]
        self.hover_bg = style_config["hover"]

        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)

    def on_enter(self, e):
        self.config(bg=self.hover_bg)

    def on_leave(self, e):
        self.config(bg=self.default_bg)


class LoadingDialog(tk.Toplevel):
    """Modern loading dialog"""

    def __init__(self, parent, message="Processing..."):
        super().__init__(parent)
        self.title("Processing")
        self.geometry("300x120")
        self.resizable(False, False)
        self.configure(bg=COLORS['white'])

        # Center the dialog
        self.transient(parent)
        self.grab_set()

        # Main frame
        main_frame = tk.Frame(self, bg=COLORS['white'])
        main_frame.pack(expand=True, fill="both", padx=20, pady=20)

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.pack(pady=(0, 10), fill="x")
        self.progress.start()

        # Message label
        tk.Label(
            main_frame,
            text=message,
            font=("Segoe UI", 10),
            bg=COLORS['white'],
            fg=COLORS['text_dark']
        ).pack()

        self.center_on_parent(parent)

    def center_on_parent(self, parent):
        parent.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - 150
        y = parent.winfo_y() + (parent.winfo_height() // 2) - 60
        self.geometry(f"300x120+{x}+{y}")


class BaseUpdateFrame(tk.Frame):
    """Enhanced base frame for data entry grids"""

    def __init__(self, master, headers, module_name):
        super().__init__(master, bg=COLORS['light_bg'])
        self.headers = headers
        self.module_name = module_name
        self.max_rows = 50
        self.num_cols = len(headers)
        self.entry_cells = []

        # Header
        header_frame = tk.Frame(self, bg=COLORS['primary'], height=60)
        header_frame.pack(fill="x", padx=0, pady=(0, 10))
        header_frame.pack_propagate(False)

        tk.Label(
            header_frame,
            text=f"{module_name}",
            font=("Segoe UI", 16, "bold"),
            bg=COLORS['primary'],
            fg=COLORS['white']
        ).pack(expand=True)

        # Main content frame
        content_frame = tk.Frame(self, bg=COLORS['light_bg'])
        content_frame.pack(fill="both", expand=True, padx=10)

        # Instructions
        instruction_frame = tk.Frame(content_frame, bg=COLORS['light_bg'])
        instruction_frame.pack(fill="x", pady=(0, 10))

        tk.Label(
            instruction_frame,
            text="Click on any cell and press Ctrl+V to paste data from Excel",
            font=("Segoe UI", 9),
            bg=COLORS['light_bg'],
            fg=COLORS['text_light']
        ).pack(side="left")

        # Action buttons
        btn_frame = tk.Frame(instruction_frame, bg=COLORS['light_bg'])
        btn_frame.pack(side="right")

        ModernButton(btn_frame, "Reset", style="warning", command=self.reset_table).pack(side="right", padx=(5, 0))
        ModernButton(btn_frame, "Save to Database", style="success", command=self.save_data).pack(side="right")

        # Create scrollable table
        self.create_table(content_frame)

        # Bind paste event
        self.bind_all("<Control-v>", self.paste_data)

    def create_table(self, parent):
        """Create scrollable data entry table"""
        # Create main container
        table_container = tk.Frame(parent, bg=COLORS['white'], relief="solid", bd=1)
        table_container.pack(fill="both", expand=True)

        # Create canvas and scrollbars
        canvas = tk.Canvas(table_container, bg=COLORS['white'])
        v_scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(table_container, orient="horizontal", command=canvas.xview)

        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # Pack scrollbars and canvas
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")
        canvas.pack(side="left", fill="both", expand=True)

        # Create table frame inside canvas
        table_frame = tk.Frame(canvas, bg=COLORS['white'])
        canvas.create_window((0, 0), window=table_frame, anchor='nw')

        # Create headers
        for col, header in enumerate(self.headers):
            header_label = tk.Label(
                table_frame,
                text=header,
                width=15,
                bg=COLORS['primary'],
                fg=COLORS['white'],
                font=("Segoe UI", 9, "bold"),
                relief="solid",
                bd=1
            )
            header_label.grid(row=0, column=col, sticky="ew")

        # Create entry cells
        for row in range(1, self.max_rows + 1):
            row_entries = []
            for col in range(self.num_cols):
                entry = tk.Entry(
                    table_frame,
                    width=15,
                    font=("Segoe UI", 9),
                    relief="solid",
                    bd=1,
                    bg=COLORS['white']
                )
                entry.grid(row=row, column=col, padx=0, pady=0, sticky="ew")

                # Alternate row colors
                if row % 2 == 0:
                    entry.configure(bg="#F8F9FA")

                row_entries.append(entry)
            self.entry_cells.append(row_entries)

        # Configure grid weights
        for col in range(self.num_cols):
            table_frame.columnconfigure(col, weight=1)

        # Update scroll region
        table_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    def paste_data(self, event=None):
        """Silent paste functionality - no success message"""
        try:
            focused_widget = self.focus_get()
            start_row = start_col = None

            # Find the focused cell
            for r_idx, row_entries in enumerate(self.entry_cells):
                for c_idx, entry in enumerate(row_entries):
                    if entry == focused_widget:
                        start_row, start_col = r_idx, c_idx
                        break
                if start_row is not None:
                    break

            if start_row is None:
                messagebox.showwarning("Paste Error", "Please click on a cell before pasting data.")
                return

            clipboard = pyperclip.paste()
            if not clipboard.strip():
                messagebox.showwarning("Paste Error", "Clipboard is empty.")
                return

            rows = clipboard.strip().split("\n")

            for i, line in enumerate(rows):
                if start_row + i >= len(self.entry_cells):
                    break

                cols = line.split("\t")
                for j, val in enumerate(cols):
                    if start_col + j < self.num_cols:
                        entry = self.entry_cells[start_row + i][start_col + j]
                        entry.delete(0, tk.END)
                        entry.insert(0, val.strip())

            # Silent paste - no success message popup

        except Exception as e:
            messagebox.showerror("Paste Error", f"Failed to paste data: {str(e)}")

    def reset_table(self):
        """Reset all entry fields"""
        confirm = messagebox.askyesno("Confirm Reset", "Are you sure you want to clear all data?")
        if confirm:
            for row in self.entry_cells:
                for entry in row:
                    entry.delete(0, tk.END)
            messagebox.showinfo("Reset Complete", "All fields have been cleared.")

    def save_data(self):
        """Override in subclasses"""
        pass


class TaxrollValueUpdate(BaseUpdateFrame):
    def __init__(self, master):
        super().__init__(master, headers_map["Value Update"], "Value Update")

    def save_data(self):
        loading = LoadingDialog(self, "Saving value updates...")

        def save_worker():
            try:
                query = """INSERT INTO Taxroll_UpdateEntries (
                    Taxyear,CADID,AccountNumber,OCALUC,LandValue,ImprovementValue,
                    MarketValues,AssessedValues,Batch,ValueType,CreatedBy,CreatedDate
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"""

                created_by = getpass.getuser()
                created_date = datetime.now()
                inserted_count = 0

                for row in self.entry_cells:
                    raw_values = [entry.get() for entry in row]
                    if not any(raw_values):
                        continue

                    values = sanitize_row(raw_values, 10, numeric_cols={4, 5, 6, 7})
                    values.extend([created_by, created_date])

                    success = insert_to_servers(query, values)
                    if success > 0:
                        session_updates["Value Update"].append(values)
                        inserted_count += 1

                self.after(0, lambda: [
                    loading.destroy(),
                    messagebox.showinfo("Success", f"{inserted_count} value update records saved successfully!")
                ])

            except Exception as e:
                self.after(0, lambda: [
                    loading.destroy(),
                    messagebox.showerror("Save Error", f"Failed to save data: {str(e)}")
                ])

        threading.Thread(target=save_worker, daemon=True).start()


class TaxrollLUCUpdate(BaseUpdateFrame):
    def __init__(self, master):
        super().__init__(master, headers_map["LUC Update"], "LUC Update")

    def save_data(self):
        loading = LoadingDialog(self, "Saving LUC updates...")

        def save_worker():
            try:
                query = """INSERT INTO Taxroll_LUCUpdateEntries (
                    Taxyear,CADID,AccountNumber,PreviousCADLUC,Update_CADLUC,
                    Update_Taxroll_LUC,Update_OCALUC,OCA_Desc,Category,Batch,CreatedBy,CreatedDate
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"""

                created_by = getpass.getuser()
                created_date = datetime.now()
                inserted_count = 0

                for row in self.entry_cells:
                    raw_values = [entry.get() for entry in row]
                    if not any(raw_values):
                        continue

                    values = sanitize_row(raw_values, 10)
                    values.extend([created_by, created_date])

                    success = insert_to_servers(query, values)
                    if success > 0:
                        session_updates["LUC Update"].append(values)
                        inserted_count += 1

                self.after(0, lambda: [
                    loading.destroy(),
                    messagebox.showinfo("Success", f"{inserted_count} LUC update records saved successfully!")
                ])

            except Exception as e:
                self.after(0, lambda: [
                    loading.destroy(),
                    messagebox.showerror("Save Error", f"Failed to save data: {str(e)}")
                ])

        threading.Thread(target=save_worker, daemon=True).start()


class TaxrollLandsizeUpdate(BaseUpdateFrame):
    def __init__(self, master):
        super().__init__(master, headers_map["Landsize Update"], "Landsize Update")

    def save_data(self):
        loading = LoadingDialog(self, "Saving landsize updates...")

        def save_worker():
            try:
                query = """INSERT INTO Taxroll_LandsizeUpdateEntries (
                    Taxyear,CADID,AccountNumber,LandSQFT,Batch,CreatedBy,CreatedDate
                ) VALUES (?,?,?,?,?,?,?)"""

                created_by = getpass.getuser()
                created_date = datetime.now()
                inserted_count = 0

                for row in self.entry_cells:
                    raw_values = [entry.get() for entry in row]
                    if not any(raw_values):
                        continue

                    values = sanitize_row(raw_values, 5, numeric_cols={3})
                    values.extend([created_by, created_date])

                    success = insert_to_servers(query, values)
                    if success > 0:
                        session_updates["Landsize Update"].append(values)
                        inserted_count += 1

                self.after(0, lambda: [
                    loading.destroy(),
                    messagebox.showinfo("Success", f"{inserted_count} landsize update records saved successfully!")
                ])

            except Exception as e:
                self.after(0, lambda: [
                    loading.destroy(),
                    messagebox.showerror("Save Error", f"Failed to save data: {str(e)}")
                ])

        threading.Thread(target=save_worker, daemon=True).start()


class TaxrollGBAUpdate(BaseUpdateFrame):
    def __init__(self, master):
        super().__init__(master, headers_map["GBA Update"], "GBA Update")

    def save_data(self):
        loading = LoadingDialog(self, "Saving GBA updates...")

        def save_worker():
            try:
                query = """INSERT INTO Taxroll_GBAUpdateEntries (
                    Taxyear,CADID,AccountNumber,GBA,NRA,YearBuilt,Batch,Remarks,CreatedBy,CreatedDate
                ) VALUES (?,?,?,?,?,?,?,?,?,?)"""

                created_by = getpass.getuser()
                created_date = datetime.now()
                inserted_count = 0

                for row in self.entry_cells:
                    raw_values = [entry.get() for entry in row]
                    if not any(raw_values):
                        continue

                    values = sanitize_row(raw_values, 8, numeric_cols={3, 4, 5})
                    values.extend([created_by, created_date])

                    success = insert_to_servers(query, values)
                    if success > 0:
                        session_updates["GBA Update"].append(values)
                        inserted_count += 1

                self.after(0, lambda: [
                    loading.destroy(),
                    messagebox.showinfo("Success", f"{inserted_count} GBA update records saved successfully!")
                ])

            except Exception as e:
                self.after(0, lambda: [
                    loading.destroy(),
                    messagebox.showerror("Save Error", f"Failed to save data: {str(e)}")
                ])

        threading.Thread(target=save_worker, daemon=True).start()


class TaxrollInsertUpload(tk.Frame):
    """Enhanced Taxroll Insert with Excel Upload"""

    def __init__(self, master):
        super().__init__(master, bg=COLORS['light_bg'])

        # Header
        header_frame = tk.Frame(self, bg=COLORS['primary'], height=60)
        header_frame.pack(fill="x", padx=0, pady=(0, 20))
        header_frame.pack_propagate(False)

        tk.Label(
            header_frame,
            text="Taxroll Insert Upload",
            font=("Segoe UI", 16, "bold"),
            bg=COLORS['primary'],
            fg=COLORS['white']
        ).pack(expand=True)

        # Main content
        content_frame = tk.Frame(self, bg=COLORS['light_bg'])
        content_frame.pack(expand=True, fill="both", padx=40, pady=20)

        # Upload section
        upload_frame = tk.Frame(content_frame, bg=COLORS['white'], relief="solid", bd=1)
        upload_frame.pack(fill="x", pady=20)

        tk.Label(
            upload_frame,
            text="Upload Excel File for Taxroll Insert",
            font=("Segoe UI", 14, "bold"),
            bg=COLORS['white'],
            fg=COLORS['text_dark']
        ).pack(pady=20)

        tk.Label(
            upload_frame,
            text="Select an Excel file (.xlsx or .xls) with the correct column headers",
            font=("Segoe UI", 10),
            bg=COLORS['white'],
            fg=COLORS['text_light']
        ).pack(pady=(0, 20))

        ModernButton(
            upload_frame,
            "Browse and Upload File",
            style="primary",
            command=self.upload_file
        ).pack(pady=(0, 20))

        # Status frame
        self.status_frame = tk.Frame(content_frame, bg=COLORS['light_bg'])
        self.status_frame.pack(fill="x", pady=10)

        self.status_label = tk.Label(
            self.status_frame,
            text="",
            font=("Segoe UI", 10),
            bg=COLORS['light_bg'],
            fg=COLORS['text_dark']
        )
        self.status_label.pack()

    def upload_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")]
        )

        if not file_path:
            return

        self.status_label.config(text=f"Processing: {os.path.basename(file_path)}", fg=COLORS['warning'])

        loading = LoadingDialog(self, "Processing Excel file...")

        def process_file():
            try:
                df = pd.read_excel(file_path, dtype=str)
                expected_headers = headers_map["Taxroll Insert"]

                if list(df.columns) != expected_headers:
                    self.after(0, lambda: [
                        loading.destroy(),
                        messagebox.showerror("Header Mismatch",
                                             "Excel headers do not match the expected format.\n"
                                             f"Expected {len(expected_headers)} columns but got {len(df.columns)}"),
                        self.status_label.config(text="Upload failed: Header mismatch", fg=COLORS['error'])
                    ])
                    return

                df = df.fillna("")
                created_by = getpass.getuser()
                created_date = datetime.now()

                query = f"""
                    INSERT INTO Taxroll_InsertEntries (
                        {",".join(expected_headers)}, CreatedBy, CreatedDate
                    ) VALUES ({",".join(['?' for _ in range(len(expected_headers) + 2)])})
                """

                inserted_count = 0
                for _, row in df.iterrows():
                    values = [val.strip() if val else None for val in row.values] + [created_by, created_date]
                    success = insert_to_servers(query, values)
                    if success > 0:
                        session_updates["Taxroll Insert"].append(values)
                        inserted_count += 1

                self.after(0, lambda: [
                    loading.destroy(),
                    messagebox.showinfo("Upload Success", f"{inserted_count} records inserted successfully!"),
                    self.status_label.config(text=f"Success: {inserted_count} records uploaded", fg=COLORS['success'])
                ])

            except Exception as e:
                self.after(0, lambda: [
                    loading.destroy(),
                    messagebox.showerror("Upload Failed", f"Error processing file: {str(e)}"),
                    self.status_label.config(text=f"Upload failed: {str(e)}", fg=COLORS['error'])
                ])

        threading.Thread(target=process_file, daemon=True).start()


class TaxrollUpdateMenu(tk.Toplevel):
    """Enhanced main update menu"""

    def __init__(self, master):
        super().__init__(master)
        self.title("Taxroll Data Update Center")
        self.geometry("1400x800")
        self.configure(bg=COLORS['light_bg'])

        # Main container
        main_container = tk.Frame(self, bg=COLORS['light_bg'])
        main_container.pack(fill="both", expand=True, padx=10, pady=10)

        # Create sidebar
        self.create_sidebar(main_container)

        # Create content area
        self.content_frame = tk.Frame(main_container, bg=COLORS['white'])
        self.content_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))

        # Initialize with welcome screen
        self.current_frame = None
        self.show_welcome()

    def create_sidebar(self, parent):
        """Create modern sidebar with navigation"""
        sidebar = tk.Frame(parent, width=250, bg=COLORS['dark_bg'])
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        # Sidebar header
        header_frame = tk.Frame(sidebar, bg=COLORS['primary'], height=80)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)

        tk.Label(
            header_frame,
            text="Update Center",
            font=("Segoe UI", 14, "bold"),
            bg=COLORS['primary'],
            fg=COLORS['white']
        ).pack(expand=True)

        # Navigation buttons
        nav_frame = tk.Frame(sidebar, bg=COLORS['dark_bg'])
        nav_frame.pack(fill="both", expand=True, padx=10, pady=20)

        # Navigation options with icons
        nav_options = [
            ("Value Update", "💰", TaxrollValueUpdate),
            ("LUC Update", "🏷️", TaxrollLUCUpdate),
            ("Landsize Update", "📏", TaxrollLandsizeUpdate),
            ("GBA Update", "🏠", TaxrollGBAUpdate),
            ("Taxroll Insert", "📤", TaxrollInsertUpload)
        ]

        for text, icon, frame_class in nav_options:
            btn_frame = tk.Frame(nav_frame, bg=COLORS['dark_bg'])
            btn_frame.pack(fill="x", pady=5)

            nav_btn = tk.Button(
                btn_frame,
                text=f"{icon} {text}",
                font=("Segoe UI", 11),
                bg=COLORS['dark_bg'],
                fg=COLORS['white'],
                relief="flat",
                bd=0,
                padx=15,
                pady=12,
                anchor="w",
                command=lambda fc=frame_class: self.show_frame(fc),
                cursor="hand2"
            )
            nav_btn.pack(fill="x")

            # Hover effects
            def on_enter(e, btn=nav_btn):
                btn.config(bg=COLORS['primary'])

            def on_leave(e, btn=nav_btn):
                btn.config(bg=COLORS['dark_bg'])

            nav_btn.bind("<Enter>", on_enter)
            nav_btn.bind("<Leave>", on_leave)

        # Submit button at bottom
        submit_frame = tk.Frame(sidebar, bg=COLORS['dark_bg'], height=80)
        submit_frame.pack(side="bottom", fill="x", padx=10, pady=10)
        submit_frame.pack_propagate(False)

        ModernButton(
            submit_frame,
            "📧 Submit Summary",
            style="success",
            command=self.submit_summary
        ).pack(expand=True, fill="x")

    def show_welcome(self):
        """Show welcome screen"""
        if self.current_frame:
            self.current_frame.destroy()

        welcome_frame = tk.Frame(self.content_frame, bg=COLORS['white'])
        welcome_frame.pack(fill="both", expand=True, padx=40, pady=40)

        # Welcome content
        tk.Label(
            welcome_frame,
            text="Welcome to Taxroll Update Center",
            font=("Segoe UI", 20, "bold"),
            bg=COLORS['white'],
            fg=COLORS['text_dark']
        ).pack(pady=(0, 20))

        tk.Label(
            welcome_frame,
            text="Select an update type from the sidebar to begin",
            font=("Segoe UI", 12),
            bg=COLORS['white'],
            fg=COLORS['text_light']
        ).pack(pady=(0, 40))

        # Quick stats if available
        stats_frame = tk.Frame(welcome_frame, bg=COLORS['light_bg'], relief="solid", bd=1)
        stats_frame.pack(fill="x", pady=20)

        tk.Label(
            stats_frame,
            text="Session Statistics",
            font=("Segoe UI", 14, "bold"),
            bg=COLORS['light_bg'],
            fg=COLORS['text_dark']
        ).pack(pady=(10, 5))

        # Display session stats
        for module, updates in session_updates.items():
            if updates:
                tk.Label(
                    stats_frame,
                    text=f"{module}: {len(updates)} records",
                    font=("Segoe UI", 10),
                    bg=COLORS['light_bg'],
                    fg=COLORS['text_dark']
                ).pack(pady=2)

        if not any(session_updates.values()):
            tk.Label(
                stats_frame,
                text="No updates recorded in this session",
                font=("Segoe UI", 10),
                bg=COLORS['light_bg'],
                fg=COLORS['text_light']
            ).pack(pady=10)

        self.current_frame = welcome_frame

    def show_frame(self, frame_class):
        """Switch to different frame"""
        if self.current_frame:
            self.current_frame.destroy()

        self.current_frame = frame_class(self.content_frame)
        self.current_frame.pack(fill="both", expand=True)

    def submit_summary(self):
        """Enhanced submit functionality"""
        if not any(session_updates.values()):
            messagebox.showwarning("No Data", "No updates found to submit in this session.")
            return

        # Show confirmation dialog
        confirm_dialog = tk.Toplevel(self)
        confirm_dialog.title("Confirm Submission")
        confirm_dialog.geometry("400x300")
        confirm_dialog.configure(bg=COLORS['white'])
        confirm_dialog.resizable(False, False)
        confirm_dialog.transient(self)
        confirm_dialog.grab_set()

        # Center the dialog
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 200
        y = self.winfo_y() + (self.winfo_height() // 2) - 150
        confirm_dialog.geometry(f"400x300+{x}+{y}")

        tk.Label(
            confirm_dialog,
            text="Submission Summary",
            font=("Segoe UI", 14, "bold"),
            bg=COLORS['white'],
            fg=COLORS['text_dark']
        ).pack(pady=20)

        # Summary details
        summary_frame = tk.Frame(confirm_dialog, bg=COLORS['light_bg'])
        summary_frame.pack(fill="both", expand=True, padx=20, pady=10)

        for module, updates in session_updates.items():
            if updates:
                tk.Label(
                    summary_frame,
                    text=f"{module}: {len(updates)} records",
                    font=("Segoe UI", 10),
                    bg=COLORS['light_bg'],
                    fg=COLORS['text_dark']
                ).pack(pady=5)

        # Buttons
        btn_frame = tk.Frame(confirm_dialog, bg=COLORS['white'])
        btn_frame.pack(side="bottom", fill="x", padx=20, pady=20)

        ModernButton(btn_frame, "Cancel", style="secondary", command=confirm_dialog.destroy).pack(side="right",
                                                                                                  padx=(5, 0))
        ModernButton(btn_frame, "Send Email", style="success",
                     command=lambda: self.send_summary_email(confirm_dialog)).pack(side="right")

    def send_summary_email(self, dialog):
        """Send summary email"""
        dialog.destroy()
        loading = LoadingDialog(self, "Preparing and sending email...")

        def email_worker():
            try:
                # Create Excel file
                tmp_file = os.path.join(tempfile.gettempdir(), "Taxroll_Update_Report.xlsx")
                export_to_excel_multi_sheet(session_updates, tmp_file, headers_map)

                # Prepare email content
                summary_lines = []
                summary_lines.append("<h3>Taxroll Update Summary</h3>")
                summary_lines.append(
                    "<table border='1' cellspacing='0' cellpadding='8' style='border-collapse: collapse;'>")
                summary_lines.append(
                    "<tr style='background-color: #f8f9fa;'><th>Module</th><th>Records</th><th>Distinct Batches</th></tr>")

                total_batches = set()
                for module, rows in session_updates.items():
                    if not rows:
                        continue

                    headers = headers_map[module]
                    batch_index = headers.index("Batch") if "Batch" in headers else None
                    batches = set()

                    if batch_index is not None:
                        for row in rows:
                            batch_value = row[batch_index]
                            if batch_value:
                                batches.add(str(batch_value).strip())
                                total_batches.add(str(batch_value).strip())

                    summary_lines.append(f"<tr><td>{module}</td><td>{len(rows)}</td><td>{len(batches)}</td></tr>")

                summary_lines.append("</table>")
                summary_lines.append(
                    f"<p><b>Total Distinct Batches:</b> {len(total_batches)} ({', '.join(sorted(total_batches))})</p>")

                # Email subject and body
                subject = f"Taxroll Updates Summary - {len(total_batches)} Batches Processed ({', '.join(sorted(total_batches))})"

                body = f"""
                <html>
                <body style="font-family: Arial, sans-serif;">
                    <h2>Taxroll Update Submission</h2>
                    <p>User <b>{getpass.getuser()}</b> has submitted taxroll updates on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.</p>
                    {''.join(summary_lines)}
                    <p>Detailed records are available in the attached Excel file.</p>
                    <hr>
                    <p><i>This is an automated email from the Taxroll Update System.</i></p>
                </body>
                </html>
                """

                success = send_email_with_attachment(subject, body, tmp_file, html=True)

                if success:
                    self.after(0, lambda: [
                        loading.destroy(),
                        messagebox.showinfo("Success", "Summary email sent successfully!"),
                        self.clear_session_data()
                    ])
                else:
                    self.after(0, lambda: [
                        loading.destroy(),
                        messagebox.showerror("Email Error",
                                             "Failed to send summary email. Please check email configuration.")
                    ])

            except Exception as e:
                self.after(0, lambda: [
                    loading.destroy(),
                    messagebox.showerror("Email Error", f"Failed to send summary email: {str(e)}")
                ])

        threading.Thread(target=email_worker, daemon=True).start()

    def clear_session_data(self):
        """Clear session data after successful submission"""
        for module in session_updates:
            session_updates[module].clear()
        self.show_welcome()  # Refresh welcome screen


class TaxrollApp(tk.Tk):
    """Enhanced main application"""

    def __init__(self):
        super().__init__()
        self.title("Pathfinder Taxroll Management System")
        self.geometry("1000x600")
        self.configure(bg=COLORS['white'])
        self.resizable(True, True)

        # Set minimum size
        self.minsize(800, 500)

        # Initialize after window is ready
        self.after(0, self.initialize_app)

    def initialize_app(self):
        """Initialize application after window is created"""
        global session_user_authenticated, authenticated_user

        # Check for saved authentication
        saved_username = load_auth_data()
        if saved_username:
            authenticated_user = saved_username
            session_user_authenticated = True
            self.show_home()
        elif session_user_authenticated:
            self.show_home()
        else:
            self.show_login()

    def show_login(self):
        """Enhanced login screen"""
        self.clear_window()

        # Main container
        main_container = tk.Frame(self, bg=COLORS['white'])
        main_container.pack(fill="both", expand=True)

        # Left panel - Branding
        left_panel = tk.Frame(main_container, bg=COLORS['primary'])
        left_panel.place(relx=0, rely=0, relwidth=0.6, relheight=1)

        # Brand content
        brand_frame = tk.Frame(left_panel, bg=COLORS['primary'])
        brand_frame.pack(expand=True)

        tk.Label(
            brand_frame,
            text="PATHFINDER",
            font=("Segoe UI", 32, "bold"),
            bg=COLORS['primary'],
            fg=COLORS['white']
        ).pack(pady=(0, 10))

        tk.Label(
            brand_frame,
            text="Taxroll Management System",
            font=("Segoe UI", 16),
            bg=COLORS['primary'],
            fg=COLORS['white']
        ).pack(pady=(0, 20))

        tk.Label(
            brand_frame,
            text="Professional • Secure • Efficient",
            font=("Segoe UI", 12),
            bg=COLORS['primary'],
            fg="#B8D4E3"
        ).pack()

        # Right panel - Login
        right_panel = tk.Frame(main_container, bg=COLORS['light_bg'])
        right_panel.place(relx=0.6, rely=0, relwidth=0.4, relheight=1)

        # Login form
        login_frame = tk.Frame(right_panel, bg=COLORS['light_bg'])
        login_frame.pack(expand=True)

        tk.Label(
            login_frame,
            text="Welcome Back",
            font=("Segoe UI", 20, "bold"),
            bg=COLORS['light_bg'],
            fg=COLORS['text_dark']
        ).pack(pady=(0, 10))

        tk.Label(
            login_frame,
            text="Please sign in with your Windows credentials",
            font=("Segoe UI", 10),
            bg=COLORS['light_bg'],
            fg=COLORS['text_light']
        ).pack(pady=(0, 30))

        # Username field
        tk.Label(
            login_frame,
            text="Username",
            font=("Segoe UI", 10, "bold"),
            bg=COLORS['light_bg'],
            fg=COLORS['text_dark']
        ).pack(anchor="w", padx=40)

        self.username_entry = tk.Entry(
            login_frame,
            font=("Segoe UI", 12),
            width=25,
            relief="solid",
            bd=1,
            bg=COLORS['white']
        )
        self.username_entry.pack(pady=(5, 15), padx=40, fill="x")

        # Password field
        tk.Label(
            login_frame,
            text="Password",
            font=("Segoe UI", 10, "bold"),
            bg=COLORS['light_bg'],
            fg=COLORS['text_dark']
        ).pack(anchor="w", padx=40)

        self.password_entry = tk.Entry(
            login_frame,
            font=("Segoe UI", 12),
            width=25,
            show="*",
            relief="solid",
            bd=1,
            bg=COLORS['white']
        )
        self.password_entry.pack(pady=(5, 20), padx=40, fill="x")

        # Status label
        self.status_label = tk.Label(
            login_frame,
            text="",
            font=("Segoe UI", 9),
            bg=COLORS['light_bg'],
            fg=COLORS['error']
        )
        self.status_label.pack(pady=(0, 15))

        # Login button
        ModernButton(
            login_frame,
            "Sign In",
            style="primary",
            command=self.authenticate_user,
            width=20
        ).pack(pady=10)

        # Bind Enter key to login
        self.bind('<Return>', lambda e: self.authenticate_user())

        # Footer
        tk.Label(
            right_panel,
            text="© 2024 Pathfinder Business Analysis Pvt Ltd",
            font=("Segoe UI", 8),
            bg=COLORS['light_bg'],
            fg=COLORS['text_light']
        ).pack(side="bottom", pady=20)

        # Focus username field
        self.username_entry.focus()

    def authenticate_user(self):
        """Enhanced authentication with persistent login"""
        global session_user_authenticated, authenticated_user

        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()

        if not username or not password:
            self.status_label.config(text="Please enter both username and password")
            return

        # Check if already authenticated
        if authenticated_user == username:
            session_user_authenticated = True
            self.show_home()
            return

        self.status_label.config(text="Authenticating...", fg=COLORS['warning'])
        self.update_idletasks()

        def auth_worker():
            success = verify_windows_login(username, password)

            if success:
                global authenticated_user, session_user_authenticated
                authenticated_user = username
                session_user_authenticated = True

                # Save authentication data for future logins
                save_auth_data(username)

                self.after(0, self.show_home)
            else:
                self.after(0, lambda: [
                    self.status_label.config(text="Authentication failed. Please check your credentials.",
                                             fg=COLORS['error']),
                    self.password_entry.delete(0, tk.END)
                ])

        threading.Thread(target=auth_worker, daemon=True).start()

    def show_home(self):
        """Enhanced home screen"""
        self.clear_window()

        # Main container
        main_container = tk.Frame(self, bg=COLORS['light_bg'])
        main_container.pack(fill="both", expand=True)

        # Header
        header_frame = tk.Frame(main_container, bg=COLORS['primary'], height=80)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)

        # Header content
        header_content = tk.Frame(header_frame, bg=COLORS['primary'])
        header_content.pack(expand=True, fill="both", padx=30)

        tk.Label(
            header_content,
            text="Taxroll Management Dashboard",
            font=("Segoe UI", 18, "bold"),
            bg=COLORS['primary'],
            fg=COLORS['white']
        ).pack(side="left", expand=True, anchor="w")

        # User info and logout
        user_frame = tk.Frame(header_content, bg=COLORS['primary'])
        user_frame.pack(side="right")

        tk.Label(
            user_frame,
            text=f"Welcome, {authenticated_user or getpass.getuser()}",
            font=("Segoe UI", 10),
            bg=COLORS['primary'],
            fg=COLORS['white']
        ).pack(side="left", padx=(0, 20))

        ModernButton(
            user_frame,
            "Logout",
            style="error",
            command=self.logout_user
        ).pack(side="right")

        # Content area
        content_frame = tk.Frame(main_container, bg=COLORS['light_bg'])
        content_frame.pack(fill="both", expand=True, padx=40, pady=40)

        # Welcome section
        welcome_frame = tk.Frame(content_frame, bg=COLORS['white'], relief="solid", bd=1)
        welcome_frame.pack(fill="x", pady=(0, 30))

        tk.Label(
            welcome_frame,
            text="System Overview",
            font=("Segoe UI", 16, "bold"),
            bg=COLORS['white'],
            fg=COLORS['text_dark']
        ).pack(pady=20)

        # Quick actions
        actions_frame = tk.Frame(content_frame, bg=COLORS['light_bg'])
        actions_frame.pack(fill="both", expand=True)

        tk.Label(
            actions_frame,
            text="Quick Actions",
            font=("Segoe UI", 14, "bold"),
            bg=COLORS['light_bg'],
            fg=COLORS['text_dark']
        ).pack(pady=(0, 20))

        # Action buttons
        btn_frame = tk.Frame(actions_frame, bg=COLORS['light_bg'])
        btn_frame.pack()

        ModernButton(
            btn_frame,
            "📊 Open Update Center",
            style="primary",
            command=self.open_update_menu,
            width=25
        ).pack(pady=10)

        ModernButton(
            btn_frame,
            "📈 View Reports",
            style="secondary",
            width=25
        ).pack(pady=10)

    def open_update_menu(self):
        """Open the taxroll update menu"""
        TaxrollUpdateMenu(self)

    def logout_user(self):
        """Logout and return to login screen"""
        global session_user_authenticated, authenticated_user
        confirm = messagebox.askyesno("Confirm Logout", "Are you sure you want to logout?")
        if confirm:
            session_user_authenticated = False
            authenticated_user = None
            clear_auth_data()  # Clear saved authentication
            self.show_login()

    def clear_window(self):
        """Clear all widgets from window"""
        for widget in self.winfo_children():
            widget.destroy()


if __name__ == "__main__":
    app = TaxrollApp()
    app.mainloop()