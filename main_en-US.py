# Declaration
print("You are using the Job-submission-status-Check-tool made by Gloridust. \nIf you like this program, consider starring it on GitHub: https://github.com/Gloridust/Job-submission-status-Check-tool \nBefore using the software, please make sure all files are correctly placed and configured. \nIf you have any questions, you can seek help in the project's README file or issues.\n")
print(f"Current configuration is as follows:\nList spreadsheet: {excel_name}\nColumn for names: {name_column}\nColumn for submission status: {status_column}\nFile types for submission: {file_extensions}\nName part precedes: {name_is_before}")
input("After confirming the information is correct, press Enter to start")

# Configuration
excel_name = "Namelist.xlsx"
name_column = "Name"
status_column = "Submission Status"
file_extensions = ['.doc', '.docx']
name_is_before = "230"


import pandas as pd

# Read Excel file
df = pd.read_excel(excel_name)

# Initialize name dictionary
name_dic = {name: 0 for name in df[name_column]}

import os

# Initialize an empty tuple to store extracted names from the homework files
work_list = tuple()

# Iterate through all files in the current directory
for filename in os.listdir('.'):
    # Assume the file format is "NameOtherInfo.docx", we extract the name from the filename
    # Simple string splitting method used here, might need adjustment according to actual situations
    if any(filename.endswith(ext) for ext in file_extensions):  # Ensure we're dealing with document files
        name_part = filename.split(name_is_before)[0]  # Use the delimiter between name and student number for splitting
        work_list += (name_part,)  # Add the name to work_list

# Iterate through each name in name_dic
for name in name_dic.keys():
    # Check if the name is in work_list
    if name in work_list:
        name_dic[name] = 1  # Exists, update to submitted
    else:
        name_dic[name] = 0  # Does not exist, keep as not submitted

# Tracking submitted
submitted_count = 0
submitted_names = "Submitted by: "
for name, status in name_dic.items():
    if status == 1:  # Check if the status is 1
        submitted_count += 1
        submitted_names += (name + ",")  # Record
print("Number of submissions:", submitted_count)
print(submitted_names)

# Tracking not submitted
not_submitted_count = 0
not_submitted_names = "Not yet submitted by: "
for name, status in name_dic.items():
    if status == 0:  # Check if the status is 0
        not_submitted_count += 1
        not_submitted_names += (name + ",")  # Record
print("Number of missing submissions:", not_submitted_count)
print(not_submitted_names)

# Convert the submission status into "Submitted" or "Not Submitted"
df[status_column] = df[name_column].map(lambda name: 'Submitted' if name_dic[name] == 1 else 'Not Submitted')

# Write the updated DataFrame back to the Excel file, assuming you want to overwrite the original file
df.to_excel(excel_name, index=False)
print("Submission statuses have been saved to the spreadsheet")
