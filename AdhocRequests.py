import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
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



# Global session updates (dict to store per module)
session_user_authenticated = False
authenticated_user = None

session_updates = {
    "Value Update": [],
    "LUC Update": [],
    "Landsize Update": [],
    "GBA Update": [],
    "Taxroll Insert": []  # ✅ Fix added
}

# Headers for Excel report
headers_map = {
    "Value Update": ["Taxyear", "CADID", "AccountNumber", "OCALUC", "LandValue", "ImprovementValue", "MarketValues", "AssessedValues", "Batch","ValueType"],
    "LUC Update": ["Taxyear", "CADID", "AccountNumber", "PreviousCADLUC", "Update_CADLUC", "Update_Taxroll_LUC", "Update_OCALUC", "OCA_Desc","Category", "Batch"],
    "Landsize Update": ["Taxyear", "CADID", "AccountNumber", "LandSQFT", "Batch"],
    "GBA Update": ["Taxyear", "CADID", "AccountNumber", "GBA", "NRA", "YearBuilt", "Batch", "Remarks"],
    "Taxroll Insert": ["CADAccountNumber", "Year", "LegalDescription", "ParcelID", "ClassCode", "Remarks",
        "LandSize", "LandUseCode", "NeighborhoodCode", "ExemptionCode", "GBA", "NRA", "Units", "Grade",
        "StreetNumber", "StreetName", "NoticedLandValue", "NoticedImprovedValue", "NoticedTotalValue", "YearBuilt",
        "OwnerName", "OwnerAddress", "OwnerAddress2", "OwnerAddress3", "OwnerCitySt", "OwnerCity", "OwnerState",
        "OwnerZip", "OwnerZip4", "NoticedMarketValue", "Keymap", "EconomicArea", "SubDivi", "PropAddress", "PropCity",
        "PropZip", "PropZip4", "CadID", "agentcode", "Agent_Name", "IsUdiAccount", "Ag_Value", "Dump_Landusecode",
        "Township", "Noticed_Date", "Farmland_Value", "Farmbuilding_value", "Revenue", "Hotel_Classification",
        "PropertyName"
    ]
}

# Email config
SMTP_SERVER = "pfbaexch.poconnor.com"
EMAIL_FROM = "Dhanasekaranr@pathfinderanalysis.com"
EMAIL_TO = "Dhanasekaranr@pathfinderanalysis.com"

# Two DB servers
DB_CONNECTIONS = [
    r'DRIVER={ODBC Driver 13 for SQL Server};SERVER=taxrollstage-db;DATABASE=TaxrollStaging;Trusted_Connection=yes;',
    r'DRIVER={ODBC Driver 13 for SQL Server};SERVER=fileprepdb;DATABASE=TaxrollStaging;Trusted_Connection=yes;'
]

# ✅ Authentication
def verify_windows_login(username, password):
    try:
        if not password:
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
    except:
        return False

# ✅ DB Insert to Two Servers
def insert_to_servers(query, values):
    for conn_str in DB_CONNECTIONS:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute(query, values)
        conn.commit()
        cursor.close()
        conn.close()

def export_to_excel_multi_sheet(data_dict, file_path, headers_map):
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for module_name, rows in data_dict.items():
            if rows:
                headers = headers_map[module_name] + ["CreatedBy", "CreatedDate"]
                df = pd.DataFrame(rows, columns=headers)
                df.to_excel(writer, sheet_name=module_name[:31], index=False)

# ✅ Sanitize Row
def sanitize_row(row, expected_len, numeric_cols=None):
    cleaned = []
    for i, val in enumerate(row):
        val = val.strip() if val else None
        if not val:
            cleaned.append(None)
        elif numeric_cols and i in numeric_cols:
            val = re.sub(r'[^0-9.-]', '', val)  # remove special chars
            try:
                cleaned.append(float(val) if val else None)
            except ValueError:
                cleaned.append(None)
        else:
            cleaned.append(val)
    while len(cleaned) < expected_len:
        cleaned.append(None)
    return cleaned

# ✅ Export Multi-sheet Excel
def export_to_excel_multi_sheet(data_dict, file_path, headers_map):
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for module_name, rows in data_dict.items():
            if rows:
                headers = headers_map[module_name] + ["CreatedBy", "CreatedDate"]
                df = pd.DataFrame(rows, columns=headers)
                df.to_excel(writer, sheet_name=module_name[:31], index=False)

# ✅ Send Email
def send_email_with_attachment(subject, body, attachment_path, html=False):
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


# ✅ Base Frame for Grids
class BaseUpdateFrame(tk.Frame):
    def __init__(self, master, headers):
        super().__init__(master)
        self.headers = headers
        self.max_rows = 500
        self.num_cols = len(headers)
        self.entry_cells = []

        canvas = tk.Canvas(self, width=1300, height=400)
        v_scrollbar = tk.Scrollbar(self, orient="vertical", command=canvas.yview)
        h_scrollbar = tk.Scrollbar(self, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        canvas.pack(side="top", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")

        table_frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=table_frame, anchor='nw')
        table_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        for col, header in enumerate(headers):
            tk.Label(table_frame, text=header, width=18, bg="#d9ead3", relief="solid").grid(row=0, column=col)

        for row in range(1, self.max_rows + 1):
            row_entries = []
            for col in range(self.num_cols):
                entry = tk.Entry(table_frame, width=18)
                entry.grid(row=row, column=col, padx=1, pady=1)
                row_entries.append(entry)
            self.entry_cells.append(row_entries)

        self.bind_all("<Control-v>", self.paste_data)

    def paste_data(self, event=None):
        try:
            focused_widget = self.focus_get()
            start_row = start_col = None
            for r_idx, row_entries in enumerate(self.entry_cells):
                for c_idx, entry in enumerate(row_entries):
                    if entry == focused_widget:
                        start_row, start_col = r_idx, c_idx
                        break
                if start_row is not None:
                    break
            if start_row is None:
                messagebox.showerror("Paste Error", "Click on a cell before pasting.")
                return
            clipboard = pyperclip.paste()
            rows = clipboard.strip().split("\n")
            for i, line in enumerate(rows):
                if start_row + i >= len(self.entry_cells):
                    break
                cols = line.split("\t")
                for j, val in enumerate(cols):
                    if j + start_col < self.num_cols:
                        self.entry_cells[start_row + i][j + start_col].delete(0, tk.END)
                        self.entry_cells[start_row + i][j + start_col].insert(0, val.strip())
        except Exception as e:
            messagebox.showerror("Paste Error", str(e))

    def reset_table(self):
        for row in self.entry_cells:
            for entry in row:
                entry.delete(0, tk.END)

# ✅ Value Update Module
class TaxrollValueUpdate(BaseUpdateFrame):
    def __init__(self, master):
        headers = headers_map["Value Update"]
        super().__init__(master, headers)
        tk.Button(self, text="Save to DB", bg="#d0f0c0", command=self.save_data).pack(side="left", padx=20, pady=10)
        tk.Button(self, text="Reset", bg="#f4cccc", command=self.reset_table).pack(side="left", padx=20, pady=10)

    def save_data(self):
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
            values = sanitize_row(raw_values, 9, numeric_cols=[4, 5, 6, 7])
            values.extend([created_by, created_date])
            insert_to_servers(query, values)
            session_updates["Value Update"].append(values)
            inserted_count += 1
        messagebox.showinfo("Success", f"{inserted_count} Value rows inserted.")

# ✅ LUC Update Module
class TaxrollLUCUpdate(BaseUpdateFrame):
    def __init__(self, master):
        headers = headers_map["LUC Update"]
        super().__init__(master, headers)
        tk.Button(self, text="Save to DB", bg="#d0f0c0", command=self.save_data).pack(side="left", padx=20, pady=10)
        tk.Button(self, text="Reset", bg="#f4cccc", command=self.reset_table).pack(side="left", padx=20, pady=10)

    def save_data(self):
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
            insert_to_servers(query, values)
            session_updates["LUC Update"].append(values)
            inserted_count += 1
        messagebox.showinfo("Success", f"{inserted_count} LUC rows inserted.")

# ✅ Landsize Update Module
class TaxrollLandsizeUpdate(BaseUpdateFrame):
    def __init__(self, master):
        headers = headers_map["Landsize Update"]
        super().__init__(master, headers)
        tk.Button(self, text="Save to DB", bg="#d0f0c0", command=self.save_data).pack(side="left", padx=20, pady=10)
        tk.Button(self, text="Reset", bg="#f4cccc", command=self.reset_table).pack(side="left", padx=20, pady=10)

    def save_data(self):
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
            values = sanitize_row(raw_values, 5, numeric_cols=[3])
            values.extend([created_by, created_date])
            insert_to_servers(query, values)
            session_updates["Landsize Update"].append(values)
            inserted_count += 1
        messagebox.showinfo("Success", f"{inserted_count} Landsize rows inserted.")

# ✅ GBA Update Module
class TaxrollGBAUpdate(BaseUpdateFrame):
    def __init__(self, master):
        headers = headers_map["GBA Update"]
        super().__init__(master, headers)
        tk.Button(self, text="Save to DB", bg="#d0f0c0", command=self.save_data).pack(side="left", padx=20, pady=10)
        tk.Button(self, text="Reset", bg="#f4cccc", command=self.reset_table).pack(side="left", padx=20, pady=10)

    def save_data(self):
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
            values = sanitize_row(raw_values, 8, numeric_cols=[3, 4, 5])
            values.extend([created_by, created_date])
            insert_to_servers(query, values)
            session_updates["GBA Update"].append(values)
            inserted_count += 1
        messagebox.showinfo("Success", f"{inserted_count} GBA rows inserted.")

class TaxrollInsertUpload(tk.Frame):
    """New Module for Taxroll Insert with Excel Upload"""
    def __init__(self, master):
        super().__init__(master)
        tk.Label(self, text="Upload Taxroll Insert Excel File", font=("Arial", 14, "bold")).pack(pady=20)
        tk.Button(self, text="Browse and Upload", width=25, bg="#d0f0c0", command=self.upload_file).pack(pady=20)

    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path, dtype=str)
            expected_headers = headers_map["Taxroll Insert"]
            if list(df.columns) != expected_headers:
                messagebox.showerror("Header Mismatch", "Excel headers do not match the expected format.")
                return

            df = df.fillna("")
            created_by = getpass.getuser()
            created_date = datetime.now()

            query = f"""
                INSERT INTO Taxroll_InsertEntries (
                    {",".join(expected_headers)}, CreatedBy, CreatedDate
                ) VALUES ({",".join(['?' for _ in range(len(expected_headers)+2)])})
            """

            inserted_count = 0
            for _, row in df.iterrows():
                values = [val.strip() if val else None for val in row.values] + [created_by, created_date]
                insert_to_servers(query, values)
                session_updates["Taxroll Insert"].append(values)
                inserted_count += 1

            messagebox.showinfo("Upload Success", f"{inserted_count} rows inserted successfully!")
        except Exception as e:
            messagebox.showerror("Upload Failed", str(e))

# ✅ Update Menu
class TaxrollUpdateMenu(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Taxroll Data Update Menu")
        self.geometry("1300x700")



        sidebar = tk.Frame(self, width=200, bg="#ececec")
        sidebar.pack(side="left", fill="y")

        content_frame = tk.Frame(self)
        content_frame.pack(side="right", fill="both", expand=True)

        tk.Label(sidebar, text="Update Options", bg="#ececec", font=("Arial", 12, "bold")).pack(pady=15)
        tk.Button(sidebar, text="Value Update", width=20, command=lambda: self.show_frame(TaxrollValueUpdate)).pack(pady=10)
        tk.Button(sidebar, text="LUC Update", width=20, command=lambda: self.show_frame(TaxrollLUCUpdate)).pack(pady=10)
        tk.Button(sidebar, text="Landsize Update", width=20, command=lambda: self.show_frame(TaxrollLandsizeUpdate)).pack(pady=10)
        tk.Button(sidebar, text="GBA Yearbuilt Update", width=20, command=lambda: self.show_frame(TaxrollGBAUpdate)).pack(pady=10)
        tk.Button(sidebar, text="Taxroll Insert", width=20, command=lambda: self.show_frame(TaxrollInsertUpload)).pack(pady=10)

        tk.Button(sidebar, text="Submit", width=20, bg="#4CAF50", fg="white", command=self.submit_summary).pack(side="bottom", pady=30)

        self.current_frame = None
        self.content_frame = content_frame

    def show_frame(self, frame_class):
        if self.current_frame:
            self.current_frame.destroy()
        self.current_frame = frame_class(self.content_frame)
        self.current_frame.pack(fill="both", expand=True)

    def submit_summary(self):
        if not any(session_updates.values()):
            messagebox.showwarning("No Data", "No updates found to submit.")
            return

        confirm = messagebox.askyesno("Confirm Submit", "Send summary email with all module data?")
        if confirm:
            try:
                # Create Excel file
                tmp_file = os.path.join(tempfile.gettempdir(), "Taxroll_Update_Report.xlsx")
                export_to_excel_multi_sheet(session_updates, tmp_file, headers_map)

                # Calculate distinct batch counts and prepare summary table
                summary_lines = []
                summary_lines.append("<h3>Taxroll Update Summary</h3>")
                summary_lines.append("<table border='1' cellspacing='0' cellpadding='5'>")
                summary_lines.append("<tr><th>Module</th><th>Records</th><th>Distinct Batches</th></tr>")

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

                # Email subject with batch count and batch IDs
                subject = f"Taxroll Updates Summary - {len(total_batches)} Batches Processed ({', '.join(sorted(total_batches))})"

                # Email body with full HTML formatting
                body = f"""
                <html>
                <body>
                    <p>User <b>{getpass.getuser()}</b> has submitted Taxroll updates.</p>
                    <p>Please find the details below:</p>
                    {''.join(summary_lines)}
                    <p>Full details are attached in the Excel file.</p>
                    <p>Regards,</p>
                    <p>Taxroll Automation Bot</p>

                </body>
                </html>
                """

                send_email_with_attachment(subject, body, tmp_file, html=True)
                messagebox.showinfo("Success", "Summary email sent successfully!")
            except Exception as e:
                messagebox.showerror("Email Error", f"Failed to send summary email: {e}")


# ✅ Main App with Login

class TaxrollApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Pathfinder | Taxroll Login")
        self.geometry("900x500")
        self.resizable(False, False)
        self.configure(bg="white")
        self.after(0, self.lazy_start)

    def lazy_start(self):
        global session_user_authenticated, authenticated_user
        if session_user_authenticated:
            self.show_home()
        else:
            self.show_login()

    def show_login(self):
        global session_user_authenticated
        self.clear_window()

        wrapper = tk.Frame(self, bg="white")
        wrapper.pack(expand=True, fill="both")

        # Left Branding Panel
        left = tk.Frame(wrapper, bg="white")
        left.place(relx=0, rely=0, relwidth=0.55, relheight=1)

        branding_img = tk.PhotoImage(file="assets/pathfinder_banner2.png")
        img_label = tk.Label(left, image=branding_img, bg="white")
        img_label.image = branding_img
        img_label.pack(expand=True, fill="both")

        # Right Login Panel
        right = tk.Frame(wrapper, bg="#f9f9f9")
        right.place(relx=0.55, rely=0, relwidth=0.45, relheight=1)

        tk.Label(right, text="Welcome to Taxroll Update", font=("Arial", 16, "bold"), fg="#333", bg="#f9f9f9").pack(pady=30)
        tk.Label(right, text="Windows Login", font=("Arial", 12), fg="#666", bg="#f9f9f9").pack(pady=(0, 20))

        # Username
        self.email_entry = tk.Entry(right, font=("Arial", 12), width=30, bd=2, relief="groove", fg="#333")
        self.email_entry.insert(0, "Username")
        self.email_entry.pack(pady=10)

        # Password
        self.password_entry = tk.Entry(right, font=("Arial", 12), width=30, bd=2, relief="groove", fg="#333", show="*")
        self.password_entry.insert(0, "Password")
        self.password_entry.pack(pady=10)

        self.status_label = tk.Label(right, text="", fg="red", bg="#f9f9f9", font=("Arial", 9))
        self.status_label.pack(pady=5)

        login_btn = tk.Button(right, text="Log In", font=("Arial", 12, "bold"), bg="#1f92f4", fg="white", relief="flat", padx=15, pady=6, command=self.login_user, activebackground="#156cc1")
        login_btn.pack(pady=15)
        login_btn.bind("<Enter>", lambda e: login_btn.config(bg="#156cc1"))
        login_btn.bind("<Leave>", lambda e: login_btn.config(bg="#1f92f4"))

        tk.Label(right, text="© Pathfinder Business Analysis (P) Ltd", font=("Arial", 8), fg="#888", bg="#f9f9f9").pack(side="bottom", pady=10)

    def login_user(self):
        global session_user_authenticated, authenticated_user

        username = self.email_entry.get().strip()
        password = self.password_entry.get().strip()

        if authenticated_user == username:
            session_user_authenticated = True
            self.show_home()
            return

        self.status_label.config(text="Authenticating... Please wait")
        self.update_idletasks()

        def validate():
            success = verify_windows_login(username, password)
            if success:
                session_user_authenticated = True
                authenticated_user = username
                self.after(0, self.show_home)
            else:
                self.after(0, lambda: messagebox.showerror("Login Failed", "Invalid credentials"))
            self.after(0, self.clear_status_label)

        threading.Thread(target=validate, daemon=True).start()

    def logout_user(self):
        global session_user_authenticated, authenticated_user
        session_user_authenticated = False
        authenticated_user = None
        self.show_login()

    def clear_status_label(self):
        if self.status_label.winfo_exists():
            self.status_label.config(text="")

    def show_home(self):
        self.clear_window()

        menu_frame = tk.Frame(self, width=1300, height=700)
        menu_frame.pack(pady=20, fill="both", expand=True)

        # ✅ Set background image
        from PIL import ImageTk, Image
        bg_image = Image.open("C:/Users/dhanasekaranr/Geocodescraping/pythonProject1/assets/menu_frame_background_resized1.png")
        bg_photo = ImageTk.PhotoImage(bg_image)

        bg_label = tk.Label(menu_frame, image=bg_photo)
        bg_label.image = bg_photo  # prevent garbage collection
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)

        # ✅ Add foreground widgets on top
        tk.Label(menu_frame, text="Taxroll Update Menu", font=("Arial", 14), bg="white").pack(pady=10)
        tk.Button(menu_frame, text="1. Taxroll Data Update", width=30, command=self.open_update_menu).pack(pady=10)
        tk.Button(menu_frame, text="Logout", width=15, bg="#f44336", fg="white", command=self.show_login).pack(pady=20)


    def open_update_menu(self):
        TaxrollUpdateMenu(self)

    def clear_window(self):
        for widget in self.winfo_children():
            widget.destroy()

if __name__ == "__main__":
    app = TaxrollApp()
    app.mainloop()
