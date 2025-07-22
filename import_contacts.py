import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import requests

# Default Bitrix24 Contact Fields
def fetch_bitrix_fields(webhook_url):
    response = requests.get(f"{webhook_url.rstrip('/')}/crm.contact.fields.json")
    result = response.json()
    if not result.get('result'):
        messagebox.showerror("Error", "Failed to fetch fields from Bitrix24.")
        return []

    fields = result['result']
    # Optional: Filter only writeable fields
    allowed_types = ['string', 'integer', 'double', 'boolean', 'enumeration', 'date', 'datetime']
    bitrix_fields = [key for key, val in fields.items() if val.get('type') in allowed_types and not val.get('isReadOnly', False)]
    field_labels = {key: val.get('title', key) for key in bitrix_fields}
    return bitrix_fields, field_labels


# 1. GUI to select Excel file
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        headers = [cell.value for cell in sheet[1]]
        map_fields(file_path, headers)

# 2. GUI to map Excel headers to Bitrix fields
def map_fields(file_path, headers):
    webhook = simple_input("Enter your Bitrix24 Webhook URL:")
    field_keys, field_labels = fetch_bitrix_fields(webhook)

    def submit_mappings():
        mappings = {}
        for i, header in enumerate(headers):
            field = combo_vars[i].get()
            if field:
                mappings[header] = field
        window.destroy()
        run_import(file_path, mappings, webhook)

    window = tk.Tk()
    window.title("Field Mapping")
    tk.Label(window, text="Map Excel Columns to Bitrix24 Contact Fields").grid(row=0, column=0, columnspan=2)

    combo_vars = []
    for i, header in enumerate(headers):
        tk.Label(window, text=header).grid(row=i+1, column=0)
        var = tk.StringVar(window)
        dropdown = tk.OptionMenu(window, var, *[""] + [f"{f} - {field_labels[f]}" for f in field_keys])
        dropdown.grid(row=i+1, column=1)
        combo_vars.append(var)

    tk.Button(window, text="Start Import", command=submit_mappings).grid(row=len(headers)+1, column=0, columnspan=2)
    window.mainloop()


# 3. Import contacts based on mappings
import os

def run_import(file_path, mappings, webhook):
    if not webhook.endswith("/"):
        webhook += "/"

    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    headers = [cell.value for cell in sheet[1]]

    # Add a new column: "BITRIX_ID"
    bitrix_id_col = len(headers)
    sheet.cell(row=1, column=bitrix_id_col + 1).value = "BITRIX_ID"

    success_count = 0
    fail_count = 0

    for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        contact_data = {"fields": {}}
        for excel_col, bitrix_field_full in mappings.items():
            bitrix_field = bitrix_field_full.split(" - ")[0]  # Extract field name only
            value = row[headers.index(excel_col)]
            if value:
                if bitrix_field in ["EMAIL", "PHONE"]:
                    contact_data["fields"][bitrix_field] = [{"VALUE": value, "VALUE_TYPE": "WORK"}]
                else:
                    contact_data["fields"][bitrix_field] = value

        try:
            r = requests.post(f"{webhook}crm.contact.add.json", json=contact_data)
            r_json = r.json()
            new_id = r_json.get("result")
            if new_id:
                success_count += 1
                sheet.cell(row=i, column=bitrix_id_col + 1).value = new_id
            else:
                fail_count += 1
        except Exception as e:
            print(f"Error on row {i}: {e}")
            fail_count += 1

    # Save the new version of the Excel
    dir_name, base_name = os.path.split(file_path)
    name_only, ext = os.path.splitext(base_name)
    new_file = os.path.join(dir_name, f"{name_only}_bitrix_imported{ext}")
    wb.save(new_file)

    messagebox.showinfo("Import Complete", f"✅ Success: {success_count}\n❌ Failed: {fail_count}\n💾 Saved: {new_file}")


# 4. Simple text input popup
def simple_input(prompt_text):
    def on_submit():
        nonlocal user_input
        user_input = entry.get()
        input_window.destroy()

    user_input = None
    input_window = tk.Tk()
    input_window.title("Input Required")
    tk.Label(input_window, text=prompt_text).pack()
    entry = tk.Entry(input_window, width=50)
    entry.pack()
    tk.Button(input_window, text="Submit", command=on_submit).pack()
    input_window.mainloop()
    return user_input

# Start the GUI
root = tk.Tk()
root.withdraw()
select_file()
