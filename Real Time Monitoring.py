import tkinter as tk
from tkinter import ttk, messagebox
import csv
from datetime import datetime
import logging

# Set up logging
log_file_path = 'application_log.txt'
logging.basicConfig(filename=log_file_path, level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Log application start
logging.info("Application started")

# Automatically select CSV file based on the current day
current_date = datetime.now().strftime('%d-%m-%Y')
csv_file_path = f'D:\\DEC-2022\\SI850RECORD\\TEST{current_date}.csv'

# Function to validate credentials
def validate_credentials():
    username = entry_username.get()
    password = entry_password.get()

    # Log the validation attempt
    logging.info(f"Validation attempt for username: {username}")

    # Check if both username and password are entered
    if username and password:
        # Check if the entered username and password are valid
        if username == 'admin' and password == 'admin':
            logging.info("Successful login")
            open_data_log_window()
        else:
            logging.warning("Invalid credentials")
            messagebox.showerror("Validation", "Invalid credentials")
    else:
        logging.warning("Incomplete credentials")
        messagebox.showerror("Validation", "Please enter both username and password")

# Function to open the data log window
def open_data_log_window():
    # Destroy the main login window
    root.destroy()

    # Log CSV file selection
    logging.info(f"CSV file selected: {csv_file_path}")

    # Create a new window with the main application as its parent
    data_log_window = tk.Tk()
    data_log_window.title("View Data Log")

    # Create a Treeview for displaying data log
    tree = ttk.Treeview(data_log_window, columns=('Machine No', 'Sl No', 'Product Name', 'Date', 'Time', 'Tare Wt', 'Net Wt', 'Gross Wt', 'Quantity', 'Remarks'), show='headings')

    # Add column headings
    tree.heading('Machine No', text='Machine No')
    tree.heading('Sl No', text='Sl No')
    tree.heading('Product Name', text='Product Name')
    tree.heading('Date', text='Date')
    tree.heading('Time', text='Time')
    tree.heading('Tare Wt', text='Tare Wt')
    tree.heading('Net Wt', text='Net Wt')
    tree.heading('Gross Wt', text='Gross Wt')
    tree.heading('Quantity', text='Quantity')
    tree.heading('Remarks', text='Remarks')

    # Set column widths
    column_widths = {'Machine No': 100, 'Sl No': 50, 'Date': 80, 'Time': 80, 'Tare Wt': 80, 'Net Wt': 80, 'Gross Wt': 80, 'Quantity': 80, 'Remarks': 100}
    for column, width in column_widths.items():
        tree.column(column, width=width, anchor=tk.CENTER)

    # Pack the treeview widget
    tree.pack(expand=True, fill='both')

    # Create a vertical scrollbar
    scrollbar = ttk.Scrollbar(data_log_window, orient='vertical', command=tree.yview)
    scrollbar.pack(side='right', fill='y')

    # Configure the Treeview to use the scrollbar
    tree.configure(yscrollcommand=scrollbar.set)

    # Define tag configurations for color-coding
    tree.tag_configure('green', background='green')
    tree.tag_configure('red', background='red')
    tree.tag_configure('yellow', background='yellow')
    tree.tag_configure('black', background='black')

    # Function to color-code the Treeview rows based on the 'Remark' field
    def color_code_row(row):
        remark = row['Remark09'].lower()
        if remark == 'ok':
            return 'green'
        elif remark == 'high':
            return 'red'
        elif remark == 'low':
            return 'yellow'
        else:
            return 'black'  # Default color

    # Function to periodically update the Treeview with new data
    def update_treeview():
        nonlocal tree

        # Clear existing data
        tree.delete(*tree.get_children())

        # Read the CSV file and update the Treeview
        with open(csv_file_path, 'r') as csv_file:
            csv_reader = csv.DictReader(csv_file)

            for row in csv_reader:
                data = (
                    row['Machine No'], row['Sl.No01'], row['Product Name02'],
                    row['Date03'], row['Time04'], row['Tare Wt05'], row['Net Wt06'],
                    row['Gross Wt07'], row['Quantity08'], row['Remark09']
                )
                color = color_code_row(row)
                tree.insert("", "end", values=data, tags=(color,))


        # Scroll to the bottom
        tree.yview_moveto(1.0)

        # Schedule the update after a certain interval (e.g., 1000 milliseconds)
        data_log_window.after(1000, update_treeview)

    # Start the update process
    update_treeview()

    # Create a button to close the data log window
    btn_close = tk.Button(data_log_window, text="Close", command=data_log_window.destroy)
    logging.info("Application Closed")
    btn_close.pack(pady=10)

# Create the main window
root = tk.Tk()
root.title("Login")

# Create and place the username label and entry field
label_username = tk.Label(root, text="Username:")
label_username.grid(row=0, column=0, padx=10, pady=10, sticky=tk.E)

entry_username = tk.Entry(root)
entry_username.grid(row=0, column=1, padx=10, pady=10)

# Create and place the password label and entry field
label_password = tk.Label(root, text="Password:")
label_password.grid(row=1, column=0, padx=10, pady=10, sticky=tk.E)

entry_password = tk.Entry(root, show="*")  # Show asterisks instead of the actual characters
entry_password.grid(row=1, column=1, padx=10, pady=10)

# Create and place the login button
btn_login = tk.Button(root, text="Login", command=validate_credentials)
btn_login.grid(row=2, column=0, columnspan=2, pady=10)

# Run the main event loop
root.mainloop()
