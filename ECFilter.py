import xlrd
import xlwt
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import requests
import shutil
import tempfile
from subprocess import Popen
import sys
from packaging.version import parse

# Version of the script
__version__ = '1.1'

# Create Tkinter root window for file selection and update functionality
root = tk.Tk()
root.title("EC Filter")
root.geometry('400x90')  # Set the window size

def perform_update():
    # URL to the latest version of the script file on GitHub
    script_url = 'https://raw.githubusercontent.com/rpaGrafit/excel-cim-filter/main/ECFilter.py'

    try:
        # Download the new script to a temporary file
        response = requests.get(script_url)
        response.raise_for_status()  # Ensure we got a successful response

        # Create a temporary file to hold the new script
        fd, temp_file_path = tempfile.mkstemp()
        with os.fdopen(fd, 'wb') as temp_file:
            temp_file.write(response.content)

        # Replace the current script with the downloaded one
        current_script_path = os.path.abspath(sys.argv[0])
        shutil.move(temp_file_path, current_script_path)

        # Now restart the script
        python = sys.executable
        os.execl(python, python, *sys.argv)
        
    except Exception as e:
        messagebox.showerror("Update Failed", f"Failed to download the latest version of the script: {e}")


# Function to check for updates
def check_for_update():
    # URL to the raw version.txt file on GitHub
    version_url = 'https://raw.githubusercontent.com/rpaGrafit/excel-cim-filter/main/version.txt'

    try:
        response = requests.get(version_url)
        response.raise_for_status()  # Ensure we got a successful response

        latest_version = response.text.strip()
        if parse(latest_version) > parse(__version__):
            response = messagebox.askyesno("Update Available", f"A newer version {latest_version} is available. Do you want to update now?")
            if response:
                perform_update()
        else:
            messagebox.showinfo("No Update Required", "You are using the latest version.")
    except requests.exceptions.RequestException as e:
        messagebox.showerror("Update Error", f"An error occurred while checking for updates: {e}")

# Function to process Excel files
def process_files():
    file_paths = filedialog.askopenfilenames(title='Select Input Files', filetypes=[('Excel Files', '*.xls')])
    if not file_paths:
        messagebox.showwarning("No Files Selected", "No files selected. Exiting.")
        return

    for input_path in file_paths:
        book = xlrd.open_workbook(input_path)
        sheet = book.sheet_by_index(0)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        headers = data[0]
        rows = data[1:]

        sums = {}
        for row in rows:
            employee = row[headers.index('Employee')]
            nr_of_companies = row[headers.index('Nr Of Companies')]
            sums[employee] = sums.get(employee, 0) + nr_of_companies

        sorted_sums = sorted(sums.items(), key=lambda x: x[0])
        input_filename = os.path.basename(input_path)
        output_filename = f'sums_{input_filename}'
        output_path = os.path.join(os.path.dirname(input_path), output_filename)
        book_out = xlwt.Workbook()
        sheet_out = book_out.add_sheet('Sums')
        sheet_out.write(0, 0, 'Employee')
        sheet_out.write(0, 1, 'Nr Of Companies')
        for i, (employee, total) in enumerate(sorted_sums):
            sheet_out.write(i+1, 0, employee)
            sheet_out.write(i+1, 1, total)
        book_out.save(output_path)

        messagebox.showinfo("File Processed", f"Saved output to {output_path}")

# Buttons for processing files and checking updates
file_button = tk.Button(root, text="Process Excel Files", command=process_files)
update_button = tk.Button(root, text="Update to Latest Version", command=check_for_update)

file_button.pack(side=tk.TOP, pady=10)
update_button.pack(side=tk.BOTTOM, pady=10)  # This will pack the button at the bottom

# Configure the update button size and padding
update_button.config(width=20, pady=5)

# Start the GUI event loop
root.mainloop()
