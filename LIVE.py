import pyodbc
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox

# Database configuration
server = r'DESKTOP-3F0F0E6\SQLEXPRESS'
database = 'SI850'
mac_settings_table = 'MACSettings'
mac_data_prefix = 'Mac_Rec'
mac_no_column_mac_settings = 'MacNo'
mac_no_column_mac_rec = 'MNo'

connection_string = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database}'
processed_records = set()

# Function to fetch and display records
def fetch_and_display_records(treeview):
    try:
        connection = pyodbc.connect(connection_string)
        cursor = connection.cursor()

        # Fetch MAC settings values
        query_mac_settings = f"SELECT DISTINCT {mac_no_column_mac_settings} FROM {mac_settings_table}"
        cursor.execute(query_mac_settings)
        mac_settings_values = [row.MacNo for row in cursor.fetchall()]

        # Fetch table names
        query_table_names = f"SELECT name FROM sys.tables WHERE name LIKE '{mac_data_prefix}%'"
        cursor.execute(query_table_names)
        table_names = [row.name for row in cursor.fetchall()]

        all_records = []  # To store all records for sorting
        time_column_index = None  # To store the index of the time column
        date_column_index = None  # To store the index of the date column
        remark_column_index = None  # To store the index of the REMARK column

        for table_name in table_names:
            query_mac_rec = f"SELECT DISTINCT {mac_no_column_mac_rec} FROM {table_name}"
            cursor.execute(query_mac_rec)
            mac_rec_values = [row.MNo for row in cursor.fetchall()]

            matching_values = set(mac_settings_values).intersection(mac_rec_values)

            if matching_values:
                # Get column names
                query_columns = f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = ?"
                cursor.execute(query_columns, table_name)
                column_names = [row.COLUMN_NAME for row in cursor.fetchall()]

                # Set column names dynamically in the Treeview
                if not treeview["columns"]:
                    treeview["columns"] = column_names
                    treeview["show"] = "headings"
                    for col in column_names:
                        treeview.heading(col, text=col)
                        treeview.column(col, width=120, anchor="center")

                # Find the index of the time, date, and REMARK columns
                time_column_index = next((i for i, col in enumerate(column_names) if "TIME" in col.upper()), None)
                date_column_index = next((i for i, col in enumerate(column_names) if "DATE" in col.upper()), None)
                remark_column_index = next((i for i, col in enumerate(column_names) if "REMARK" in col.upper()), None)

                # Fetch records from the table
                query_records = f"SELECT * FROM {table_name}"
                cursor.execute(query_records)
                records = cursor.fetchall()

                for record in records:
                    record_id = tuple(record)
                    if record_id not in processed_records:
                        processed_records.add(record_id)
                        all_records.append(record)

        # Sort records by date and time columns if found
        if date_column_index is not None and time_column_index is not None:
            all_records.sort(
                key=lambda x: datetime.strptime(f"{x[date_column_index]} {x[time_column_index]}", "%d-%m-%Y %H:%M:%S")
                if x[date_column_index] and x[time_column_index] else datetime.min
            )

        # Insert records into the Treeview
        for record in all_records:
            cleaned_record = [str(item).replace(",", "").replace("'", "") for item in record]
            if tuple(cleaned_record) not in processed_records:
                # Determine the tag for the record based on the REMARK column
                if remark_column_index is not None:
                    remark = cleaned_record[remark_column_index].upper()
                    if remark == "OK":
                        tag = "ok"
                    elif remark == "LOW":
                        tag = "low"
                    elif remark == "HIGH":
                        tag = "high"
                    else:
                        tag = ""
                else:
                    tag = ""

                treeview.insert("", "end", values=cleaned_record, tags=(tag,))
                processed_records.add(tuple(cleaned_record))

        # Scroll to the bottom (new record)
        treeview.yview_moveto(1)

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
    finally:
        try:
            cursor.close()
        except:
            pass
        try:
            connection.close()
        except:
            pass

# Function to handle periodic updates of the Treeview
def periodic_update(treeview, interval=5000):
    fetch_and_display_records(treeview)
    treeview.after(interval, lambda: periodic_update(treeview, interval))

# Function to handle login and open records window
def handle_login():
    username = username_entry.get()
    password = password_entry.get()

    # Replace with your desired username and password
    if username == "admin" and password == "admin":
        security_window.destroy()  # Close the login window
        open_records_window()  # Open the records display window
    else:
        messagebox.showerror("Access Denied", "Incorrect username or password. Please try again.")

# Function to create the records display window
def open_records_window():
    records_window = tk.Tk()
    records_window.title("Records Display")
    records_window.geometry("800x400")

    # Create Treeview widget
    treeview = ttk.Treeview(records_window)
    treeview.pack(fill="both", expand=True)

    # Create styles for tags
    style = ttk.Style(records_window)
    style.configure("Treeview", rowheight=25)
    style.map(
        "Treeview",
        background=[("selected", "blue")],
    )

    # Tag configurations for record colors
    treeview.tag_configure("ok", background="lightgreen")
    treeview.tag_configure("low", background="yellow")
    treeview.tag_configure("high", background="red")

    # Apply borders to rows by alternating row colors
    for i, item in enumerate(treeview.get_children()):
        if i % 2 == 0:
            treeview.item(item, tags=('even',))
        else:
            treeview.item(item, tags=('odd',))

    # Add style for row borders
    style.configure('odd', background='white')
    style.configure('even', background='lightgray')

    # Start periodic updates
    periodic_update(treeview)

    records_window.mainloop()

# Create the Tkinter security window
security_window = tk.Tk()
security_window.title("Security Check")
security_window.geometry("300x200")

# Add labels and entries for username and password
tk.Label(security_window, text="Username:", font=("Arial", 12)).pack(pady=5)
username_entry = tk.Entry(security_window, font=("Arial", 12))
username_entry.pack()

tk.Label(security_window, text="Password:", font=("Arial", 12)).pack(pady=5)
password_entry = tk.Entry(security_window, font=("Arial", 12), show="*")
password_entry.pack()

# Add login button
login_button = tk.Button(security_window, text="Login", font=("Arial", 12), command=handle_login)
login_button.pack(pady=15)

# Run the Tkinter event loop
security_window.mainloop()
