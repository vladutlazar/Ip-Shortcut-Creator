import pandas as pd
import os
import subprocess
import re

# Define a function to sanitize the PC name
def sanitize_pc_name(pc_name):
    # Remove invalid characters (except alphanumeric, underscore, hyphen, space, period, and ampersand)
    sanitized_name = re.sub(r'[^\w\s.&]', '', pc_name)
    return sanitized_name.strip()  # Remove leading/trailing spaces

# Loading the spreadsheet and limiting the range to avoid reading beyond line 155
excel_path = r'\\vt1.vitesco.com\SMT\didt1002\05_IT_MES\IP_MES.xlsx'
df = pd.read_excel(excel_path, sheet_name="IP", nrows=155)  # Read only the first 155 rows

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    pc_name = row.get('Nume Linie', 'Default Value')
    pc_name_sanitized = sanitize_pc_name(str(pc_name))  # Convert to string and sanitize
    pc_ip = row.get('IP ', 'Default Value')

    # Defining the target path and the shortcut file
    target_path = f"\\\\{pc_ip}\d$"
    shortcut_file = f"{pc_name_sanitized}.lnk"  # Use sanitized name

    # Skip the row if it contains empty or invalid data
    if pd.isna(pc_name) or pd.isna(pc_ip):
        continue

    # Defining the location where the shortcuts are created
    shortcut_folder = r'\\vt1.vitesco.com\SMT\didt1083\01_MES_PUBLIC\PC_Prod'

     # Check if the shortcut file already exists
    if not os.path.exists(os.path.join(shortcut_folder, shortcut_file)):
        try:
            # Using subprocess to run a command that creates the shortcut
            command = f'powershell "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut(\'{os.path.join(shortcut_folder, shortcut_file)}\'); $s.TargetPath = \'{target_path}\'; $s.Description = \'{pc_ip}\'; $s.Save()"'
            subprocess.run(command, shell=True)
            print(f"Shortcut for {pc_name} created successfully!")
        except Exception as e:
            print(f"Error creating shortcut for {pc_name}: {str(e)}")
    else:
        print(f"Shortcut for {pc_name} already exists. Skipping.")

print("All shortcuts processed.")