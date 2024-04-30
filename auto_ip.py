import pandas as pd
import os
import subprocess
import re

# Define a function to sanitize the PC name
def sanitize_pc_name(pc_name):
    # Remove invalid characters (except alphanumeric, underscore, hyphen, space, period, and ampersand)
    sanitized_name = re.sub(r'[^\w\s.&-]', '', pc_name)
    return sanitized_name.strip()  # Remove leading/trailing spaces

# Define a function to get the description from a shortcut
def get_shortcut_description(shortcut_path):
    # Escape single quotes within the PowerShell command
    command = f'''powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('{shortcut_path.replace("'", "''")}'); $s.Description"'''
    result = subprocess.run(command, shell=True, capture_output=True, text=True)
    return result.stdout.strip()

# Load the spreadsheet and limit the range to 155 rows
excel_path = r'\\vt1.vitesco.com\SMT\didt1002\05_IT_MES\IP_MES.xlsx'
df = pd.read_excel(excel_path, sheet_name="IP", nrows=155)

# Define the location where the shortcuts are created
shortcut_folder = r'\\vt1.vitesco.com\SMT\didt1083\01_MES_PUBLIC\1.5.PC_Prod'

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    pc_name = row.get('Nume Linie', 'Default Value')
    pc_name_sanitized = sanitize_pc_name(str(pc_name))
    pc_ip_cell = row.get('IP ', 'Default Value')

    # Handle NaN or empty cells
    if pd.isna(pc_ip_cell) or pc_ip_cell == "":
        continue

    # Extract the latest IP address and clean it
    pc_ip_list = str(pc_ip_cell).splitlines()
    pc_ip_latest = pc_ip_list[-1] if pc_ip_list else None

    if pc_ip_latest:
        pc_ip_clean = re.sub(r'[^0-9.]', '', pc_ip_latest)
    else:
        continue

    # Change the target path based on the row index
    if index >= 121:
        target_path = f"\\\\{pc_ip_clean}\\c$"
    else:
        target_path = f"\\\\{pc_ip_clean}\\d$"

    # Create the shortcut file path and escape single quotes
    shortcut_file_path = os.path.join(shortcut_folder, f"{pc_name_sanitized}.lnk").replace("'", "''")

    if os.path.exists(shortcut_file_path):
        # Check if the existing shortcut's description matches the current IP
        existing_ip = get_shortcut_description(shortcut_file_path)
        if existing_ip == pc_ip_clean:
            print(f"Shortcut for {pc_name_sanitized} already exists with the same IP. Skipping.")
            continue
        else:
            print(f"Updating shortcut for {pc_name_sanitized} to new IP {pc_ip_clean}.")
    else:
        print(f"Creating shortcut for {pc_name_sanitized}.")

    # Create or update the shortcut
    try:
        command = f'''powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('{shortcut_file_path}'); $s.TargetPath = '{target_path}'; $s.Description = '{pc_ip_clean}'; $s.Save()"'''
        subprocess.run(command, shell=True)
        print(f"Shortcut for {pc_name_sanitized} created or updated successfully!")
    except Exception as e:
        print(f"Error creating/updating shortcut for {pc_name_sanitized}: {str(e)}")

print("All shortcuts processed.")
