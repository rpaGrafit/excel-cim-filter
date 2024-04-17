import xlrd
import xlwt
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import requests

# Version of the script
__version__ = '1.0'

# Create Tkinter root window for file selection and update functionality
root = tk.Tk()
root.title("EC Filter")

# Function to check for updates
def check_for_update():
    # URL to your version file on GitHub, update it with your actual URL
    version_url = 'https://raw.githubusercontent.com/rpaGraft/excel-cim-filter/main/version.txt'
    try:
        response = requests.get(version_url)
        latest_version = response.text.strip()  # Ensure your version file contains only the version number
        if latest_version > __version__:
            messagebox.showinfo("Update Available", f"A newer version {latest_version} is available.")
            # Add further code to handle the download and update process
        else:
            messagebox.showinfo("No Update Required", "You are using the latest version.")
    except requests.exceptions.RequestException:
        messagebox.showerror("Update Error", "Failed to connect to the update server. Check your internet connection.")

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

file_button.pack(pady=10)
update_button.pack(pady=10)

# Start the GUI event loop
root.mainloop()
