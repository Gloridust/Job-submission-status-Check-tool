# Configuration Information

excel_name = "Namelist.xlsx"
name_column = "Name"
status_column = "Submission Status"
file_extensions = ['.doc', '.docx']
name_is_before="230"

import pandas as pd
import os
import keyboard

# Introduction
print("You are using the Job Submission Status Check Tool developed by Gloridust. \nIf you find this tool helpful, consider starring the project on GitHub: https://github.com/Gloridust/Job-submission-status-Check-tool \nBefore using the software, ensure that you have correctly placed all files and configured the settings properly. \nIf you encounter any issues, feel free to seek help in the project's readme file or issues section.\n")
print(f"Current configuration:\nExcel file: {excel_name}\nName column: {name_column}\nSubmission status column: {status_column}\nFile types for submission: {file_extensions}\n")
print("Press Enter to begin after verifying the configuration.")
keyboard.wait('Enter')

# Check if Excel file exists
if not os.path.isfile(excel_name):
    print(f">>>Error: Excel file '{excel_name}' not found. Please check if the configuration information is correct.")
    print("Press Enter to exit...")
    keyboard.wait('Enter')
    exit()

# Read the Excel file
try:
    df = pd.read_excel(excel_name)
except Exception as e:
    print(f">>>Error: There was an issue reading the Excel file '{excel_name}': {e}")
    print("Press Enter to exit...")
    keyboard.wait('Enter')
    exit()

# Check if the name column exists
if name_column not in df.columns:
    print(f">>>Error: Name column '{name_column}' not found in the Excel file. Please check if the configuration information is correct.")
    print("Press Enter to exit...")
    keyboard.wait('Enter')
    exit()

print("Configuration is valid.")

# Read Excel file
df = pd.read_excel(excel_name)

# Initialize name dictionary
name_dic = {name: 0 for name in df[name_column]}

# Initialize
work_names = ''
work_num = 0

# Traverse all files in the current directory and store their names
for filename in os.listdir('.'):
    if any(filename.endswith(ext) for ext in file_extensions):  
        work_names += (filename + ",")
        work_num += 1   

print(">>>Number of detected files:", work_num)

# Check if names are in work_names
for name in name_dic.keys():
    if name in work_names:
        name_dic[name] = 1
    else:
        name_dic[name] = 0

have_sub_num = 0
have_sub = ">>>Submitted individuals: "
for name, status in name_dic.items():
    if status == 1:
        have_sub_num += 1
        have_sub += (name + ",")
print(">>>Number of submissions:", have_sub_num)
print(have_sub)

not_sub_num = 0
not_sub = ">>>Individuals who haven't submitted: "
for name, status in name_dic.items():
    if status == 0:
        not_sub_num += 1
        not_sub += (name + ",")
print(">>>Number of non-submissions:", not_sub_num)
print(not_sub)

if work_num == have_sub_num:
    pass
else:
    print(">>>Warning: Detected file count does not match the identified number of submissions<<<\n>>>Please check file naming conventions<<<")

# Convert submission status to "Submitted" or "Not Submitted"
df[status_column] = df[name_column].map(lambda name: 'Submitted' if name_dic[name] == 1 else 'Not Submitted')

# Write the updated DataFrame back to the Excel file
df.to_excel(excel_name, index=False)
print(">>>Submission status has been saved to the spreadsheet.")
print("Press Enter to exit...")
keyboard.wait('Enter')