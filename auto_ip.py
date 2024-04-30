import pandas as pd
import os
import subprocess
import re

# Function to sanitize PC names and join multi-line names into one line
def sanitize_pc_name(pc_name):
    # Remove invalid characters except alphanumeric, underscore, hyphen, space, period, and ampersand
    sanitized_name = re.sub(r'[^\w\s.&-]', '', pc_name)
    # Join multi-line names into a single line
    sanitized_name_single_line = ' '.join(sanitized_name.splitlines())
    return sanitized_name_single_line.strip()  # Remove leading/trailing spaces

# Function to read the description from a shortcut
def get_shortcut_description(shortcut_path):
    command = f'''powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('{shortcut_path.replace("'", "''")}'); $s.Description"'''
    result = subprocess.run(command, shell=True, capture_output=True, text=True)
    return result.stdout.strip()

# Load the Excel data
excel_path = r'\\vt1.vitesco.com\SMT\didt1002\05_IT_MES\IP_MES.xlsx'
df = pd.read_excel(excel_path, sheet_name="IP", nrows=155)  # Read only the first 155 rows

# Define where shortcuts are created
shortcut_folder = r'\\vt1.vitesco.com\SMT\didt1083\01_MES_PUBLIC\1.5.PC_Prod'

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    pc_name = row.get('Nume Linie', 'Default Value')
    # Sanitize and join multi-line names into a single line
    pc_name_sanitized = sanitize_pc_name(str(pc_name))
    pc_ip_cell = row.get('IP ', 'Default Value')

    # Skip invalid or empty cells
    if pd.isna(pc_ip_cell) or pc_ip_cell == "":
        continue

    # Extract and clean the IP address
    pc_ip_list = str(pc_ip_cell).splitlines()
    pc_ip_latest = pc_ip_list[-1] if pc_ip_list else None

    if pc_ip_latest:
        pc_ip_clean = re.sub(r'[^0-9.]', '', pc_ip_latest)
    else:
        continue

    # Define the target path based on the row index
    target_path = f"\\\\{pc_ip_clean}\\{'c' if index >= 121 else 'd'}$"

    # Create the shortcut file path, escaping single quotes if necessary
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

    # Create or update the shortcut using a PowerShell command
    try:
        command = [
            "powershell",
            "-Command",
            (
                "$ws = New-Object -ComObject WScript.Shell; "
                f"$s = $ws.CreateShortcut('{shortcut_file_path}'); "
                f"$s.TargetPath = '{target_path}'; "
                f"$s.Description = '{pc_ip_clean}'; "
                "$s.Save()"
            )
        ]
        subprocess.run(command, shell=True)
        print(f"Shortcut for {pc_name_sanitized} created or updated successfully!")
    except Exception as e:
        print(f"Error creating/updating shortcut for {pc_name_sanitized}: {str(e)}")

print("All shortcuts processed.")
