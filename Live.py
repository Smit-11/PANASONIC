import pyodbc
import time
import tkinter as tk
from tkinter import ttk, messagebox

# ---------------- DATABASE CONFIG ----------------
server = r'DESKTOP-JHNRQ3Q\SQLEXPRESS'
database = 'SI850'

connection_string = (
    f'DRIVER={{SQL Server}};'
    f'SERVER={server};'
    f'DATABASE={database};'
    f'Trusted_Connection=yes;'
)

# ---------------- SETTINGS ----------------
OFFLINE_DELAY = 120

# ---------------- GLOBAL STORAGE ----------------
loaded_records = set()  # Track which records we've already added
machine_status = {}
machine_timer = {}
machine_last_remark = {}

# ---------------- FETCH NEW DATA ONLY ----------------
def fetch_new_data():
    connection = pyodbc.connect(connection_string)
    cursor = connection.cursor()

    cursor.execute("""
        SELECT name
        FROM sys.tables
        WHERE name LIKE 'MAC_Rec%'
    """)

    tables = [row.name for row in cursor.fetchall()]
    all_new_data = []

    for table in tables:
        try:
            # Read columns
            cursor.execute(f"SELECT TOP 1 * FROM {table}")
            columns = [col[0] for col in cursor.description]

            # Find machine column
            machine_col = None
            possible_machine_columns = ["M/C NO04", "MACHINE NAME04", "MNo"]

            for col in possible_machine_columns:
                if col in columns:
                    machine_col = col
                    break

            if not machine_col:
                continue

            # Find remark column
            remark_col = None
            for col in columns:
                if "REMARK" in col.upper():
                    remark_col = col
                    break

            # Fetch records ORDERED BY Rec_NO DESC (newest first)
            query = f"""
                SELECT *
                FROM {table}
                ORDER BY Rec_NO DESC
            """

            cursor.execute(query)
            remark_idx = remark_col if remark_col else None
            if remark_idx:
                remark_idx = columns.index(remark_col)

            rows = cursor.fetchall()

            for row in rows:
                rec_no = str(row[0])
                
                # Only process records we haven't seen before
                if rec_no not in loaded_records:
                    loaded_records.add(rec_no)
                    
                    record = tuple(str(x) for x in row)
                    machine_no = str(row[columns.index(machine_col)])
                    
                    all_new_data.append(
                        (
                            rec_no,
                            machine_no,
                            record,
                            remark_idx,
                            columns
                        )
                    )

        except Exception as e:
            print(f"ERROR in table {table}: {e}")

    cursor.close()
    connection.close()
    
    # Sort new records by Rec_NO descending (newest first)
    all_new_data.sort(key=lambda x: int(x[0]) if x[0].isdigit() else 0, reverse=True)
    
    return all_new_data

# ---------------- LIVE TAB ----------------
def update_live(tree):
    new_data = fetch_new_data()
    current_time = time.time()

    # ---------------- CREATE COLUMNS ----------------
    if not tree["columns"] and new_data:
        columns = new_data[0][4]
        tree["columns"] = ["Machine"] + columns
        tree["show"] = "headings"
        for col in tree["columns"]:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="center", stretch=False)

    # ---------------- INSERT NEW RECORDS AT THE TOP ----------------
    for rec_no, machine_no, record, remark_idx, columns in new_data:
        # Update machine status
        machine_timer[machine_no] = current_time
        machine_status[machine_no] = "Online"

        # Determine tag color
        tag = ""
        remark_value = ""
        if remark_idx is not None:
            remark_value = str(record[remark_idx]).upper()
            if remark_value == "OK":
                tag = "ok"
            elif remark_value == "LOW":
                tag = "low"
            elif remark_value == "HIGH":
                tag = "high"

        # Insert new row at the TOP (index 0)
        item = tree.insert("", 0, values=(machine_no,) + record, tags=(tag,))
        
        # Flash effect
        tree.item(item, tags=("flash",))
        tree.after(500, lambda i=item, t=tag: tree.item(i, tags=(t,)))
        
        machine_last_remark[machine_no] = remark_value
        
        print(f"Added new record: {rec_no} for machine {machine_no}")  # Debug

    tree.after(2000, lambda: update_live(tree))

# ---------------- OFFLINE CHECKER ----------------
def check_offline(root):
    current = time.time()
    for machine in list(machine_status.keys()):
        last_time = machine_timer.get(machine, 0)
        if current - last_time > OFFLINE_DELAY:
            machine_status[machine] = "Offline"
        else:
            machine_status[machine] = "Online"
    root.after(2000, lambda: check_offline(root))

# ---------------- STATUS TAB ----------------
def update_status(tree):
    tree.delete(*tree.get_children())
    for machine in sorted(machine_status.keys()):
        status = machine_status[machine]
        last_update = machine_timer.get(machine, 0)
        time_since = int(time.time() - last_update)

        if status == "Online":
            display = "🟢 ONLINE"
            tag = "online"
        else:
            display = f"🔴 OFFLINE ({time_since}s ago)"
            tag = "offline"

        if machine in machine_last_remark:
            display += f" - {machine_last_remark[machine]}"

        tree.insert("", "end", values=(machine, display), tags=(tag,))
    tree.after(2000, lambda: update_status(tree))

# ---------------- MAIN WINDOW ----------------
def open_main_window():
    root = tk.Tk()
    root.title("Machine Monitoring System")
    root.geometry("1400x700")

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    # ---------------- LIVE TAB ----------------
    live_frame = ttk.Frame(notebook)
    notebook.add(live_frame, text="Live Monitoring")

    tree_frame = ttk.Frame(live_frame)
    tree_frame.pack(fill="both", expand=True)

    # Scrollbars
    v_scroll = ttk.Scrollbar(tree_frame, orient="vertical")
    v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal")
    h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

    # Treeview
    live_tree = ttk.Treeview(
        tree_frame,
        yscrollcommand=v_scroll.set,
        xscrollcommand=h_scroll.set
    )
    live_tree.pack(fill="both", expand=True)
    v_scroll.config(command=live_tree.yview)
    h_scroll.config(command=live_tree.xview)

    # Colors
    live_tree.tag_configure("ok", background="#90EE90")
    live_tree.tag_configure("low", background="#FFD700")
    live_tree.tag_configure("high", background="#FF6B6B")
    live_tree.tag_configure("flash", background="#87CEEB")

    update_live(live_tree)

    # ---------------- STATUS TAB ----------------
    status_frame = ttk.Frame(notebook)
    notebook.add(status_frame, text="Machine Status")

    status_tree = ttk.Treeview(
        status_frame,
        columns=("Machine", "Status"),
        show="headings"
    )
    status_tree.pack(fill="both", expand=True)
    status_tree.heading("Machine", text="Machine")
    status_tree.heading("Status", text="Status")
    status_tree.column("Machine", width=250, anchor="center")
    status_tree.column("Status", width=400, anchor="center")
    
    status_tree.tag_configure("online", background="#90EE90")
    status_tree.tag_configure("offline", background="#D3D3D3")

    update_status(status_tree)
    check_offline(root)
    root.mainloop()

# ---------------- LOGIN ----------------
def handle_login():
    username = username_entry.get()
    password = password_entry.get()

    if username == "admin" and password == "admin":
        login.destroy()
        open_main_window()
    else:
        messagebox.showerror("Login Failed", "Invalid Username or Password")

# ---------------- LOGIN WINDOW ----------------
login = tk.Tk()
login.title("Login")
login.geometry("300x220")
login.resizable(False, False)

tk.Label(login, text="Username", font=("Arial", 12)).pack(pady=10)
username_entry = tk.Entry(login, font=("Arial", 12))
username_entry.pack()

tk.Label(login, text="Password", font=("Arial", 12)).pack(pady=10)
password_entry = tk.Entry(login, show="*", font=("Arial", 12))
password_entry.pack()

tk.Button(login, text="Login", font=("Arial", 12), width=15, command=handle_login).pack(pady=20)

login.mainloop()