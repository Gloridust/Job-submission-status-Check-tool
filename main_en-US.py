# Configuration Information

excel_name = "Namelist.xlsx"
name_column = "Name"
status_column = "Submission Status"
file_extensions = ['.doc', '.docx']
name_is_before="230"

import pandas as pd
import os

# Declaration
print("You are using the Job-submission-status-Check-tool created by Gloridust. \nIf you like this program, consider giving it a star on GitHub: https://github.com/Gloridust/Job-submission-status-Check-tool \nBefore using the software, please ensure that you have placed all files correctly and configured them properly. \nIf you encounter any issues, you can seek help in the readme file or issues of the project.\n")
print(f"Current configuration information:\nExcel sheet: {excel_name}\nColumn for names: {name_column}\nColumn for submission status: {status_column}\nFile extensions for submission: {file_extensions}\nPrefix for names: {name_is_before}")
input("Press Enter to begin after confirming everything is correct")

# Read Excel file
df = pd.read_excel(excel_name)

# Initialize name dictionary
name_dic = {name: 0 for name in df[name_column]}

# Initialize an empty tuple to store extracted names from homework files
work_list = tuple()
# Record the number of files
work_num = 0

# Traverse all files in the current directory
for filename in os.listdir('.'):
    # Assuming the file name format is "Name other_info.docx", we extract the name from the filename
    # Here, a simple string splitting method is used, which may need adjustments based on the actual situation
    if any(filename.endswith(ext) for ext in file_extensions):  # Ensure processing document files
        name_part = filename.split(name_is_before)[0]  # Split using the separator between name and ID
        work_list += (name_part,)  # Add the name to work_list
        work_num += 1   # Record file count

print(">>> Detected files count:", work_num)

# Traverse each name in name_dic
for name in name_dic.keys():
    # Check if the name is in work_list
    if name in work_list:
        name_dic[name] = 1  # Exists, update to submitted
    else:
        name_dic[name] = 0  # Doesn't exist, keep as not submitted

have_sub_num = 0
have_sub = ">>> Submitted individuals: "
for name, status in name_dic.items():
    if status == 1:  # Check if the value is 1
        have_sub_num += 1
        have_sub += (name + ",")  # Record
print(">>> Number of submitted individuals:", have_sub_num)
print(have_sub)

not_sub_num = 0
not_sub = ">>> Individuals yet to submit: "
for name, status in name_dic.items():
    if status == 0:  # Check if the value is 0
        not_sub_num += 1
        not_sub += (name + ",")  # Record
print(">>> Number of individuals yet to submit:", not_sub_num)
print(not_sub)

if work_num == have_sub_num:
    pass
else:
    print(">>> Attention: Detected file count does not match the identified number of submissions <<<\n>>> Please check if file naming is standardized <<<")

# Convert submission status to "Submitted" or "Not submitted"
df[status_column] = df[name_column].map(lambda name: 'Submitted' if name_dic[name] == 1 else 'Not submitted')

# Write the updated DataFrame back to the Excel file, assuming you want to keep the original file name and overwrite it
df.to_excel(excel_name, index=False)
print(">>> Submission status has been saved to the spreadsheet")
input("Press Enter to end...")
