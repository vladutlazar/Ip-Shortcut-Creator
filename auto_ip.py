import pandas as pd
import os
import subprocess

# Loading the spreadsheet
df = pd.read_excel(r'\\vt1.vitesco.com\SMT\didt1002\05_IT_MES\IP_MES.xlsx', sheet_name="IP")

# Iterate over the rows of the DataFrame
for index, row in df.iterrows():
    pc_name = row.get('Nume Linie', 'Default Value')
    pc_ip = row.get('IP', 'Default Value')
    pc_order = row.get('NR Crt', 'Default Value')

    # Defining the target path and the shortcut file
    target_path = f"\\\\{pc_order}"
    shortcut_file = f"{pc_name}.lnk"

    # Defining the location where the shortcuts are created
    shortcut_folder = r'C:\Users\uiv55706\Desktop\PC_Prod'

    # Using subprocess to run a command that creates the shortcut
    command = f'powershell "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut(\'{os.path.join(shortcut_folder, shortcut_file)}\'); $s.TargetPath = \'{target_path}\'; $s.Save()"'
    subprocess.run(command, shell=True)

print("Shortcuts created successfully!")
