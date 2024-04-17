import xlrd
import xlwt
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import requests
from packaging.version import parse

# Version of the script
__version__ = '1.0'

# Create Tkinter root window for file selection and update functionality
root = tk.Tk()
root.title("EC Filter")
root.geometry('400x90')  # Set the window size

# Function to check for updates
def check_for_update():
    # URL to the raw version.txt file on GitHub
    version_url = 'https://raw.githubusercontent.com/rpaGrafit/excel-cim-filter/main/version.txt'

    try:
        response = requests.get(version_url)
        if response.status_code == 200:
            latest_version = response.text.strip()
            if parse(latest_version) > parse(__version__):
                messagebox.showinfo("Update Available", f"A newer version {latest_version} is available.")
                # Here you can add code to download and update the script or direct the user to the download page
            else:
                messagebox.showinfo("No Update Required", "You are using the latest version.")
        else:
            messagebox.showerror("Update Error", f"Could not retrieve the latest version. The server returned status code: {response.status_code}")

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
