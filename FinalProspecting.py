import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import os
from tkinter import *
from tkinter import filedialog

# Get the path for raw data, name/directory for export file
def get_paths():
    paths = []
    
    # initialize ability to browse files
    root = Tk()
    root.withdraw()
    
    # user chooses raw data file
    print("Select raw data file:")
    root.rawdata =  filedialog.askopenfilename(initialdir = "/",title = "Select Raw Data File")
    paths.append(root.rawdata)
    print()

    # user chooses how to name final, exported file
    print("How would you like to name your finished file?")
    final_file_name = input()
    print()

    print("Select export file:")
    root.finalfile =  filedialog.askdirectory(initialdir = "/", title = "Select Export Location")
    print(root.finalfile)

    # get the final path
    final_file_name += ".xlsx"
    final_path = root.finalfile + "/" + final_file_name
    paths.append(final_path)

    return(paths)

# Get the data frame of raw data, get info, and turn it into a list
def info_data_email(raw_data_path):
    # get the raw data and read it
    path = raw_data_path

    df = pd.read_excel(path, dtype=str)

    # Leave only column with "….." dividers, remove other two
    df = df.drop(columns = ["Col", "….."])

    # Leave only the last column (has info and "…..") and make elements all strings
    df = df["Col.1"].astype(str)

    # Turn data into a list 
    info_list = df.tolist()

    return (info_list)

# Get info for each row and create list of final info
def get_rows_email(info_list):
    info = info_list
    final_list = []

    # Get index of all "….."
    indices = [i for i, x in enumerate(info) if x == "….."]

    # get the individual row info from the info
    # get the first row (fencepost)
    
    # get index of first instance of "….."
    first_index = indices[0]
    
    # use index to get info for first row from info
    row_info = info[:first_index]

    # clean the row of unnecessary info
    clean_row = clean_rows_email(row_info)

    # add row_info to final_list
    final_list.extend(clean_row)
    
    # get all the rest of the rows 
    count = 0
    while count < len(indices) - 1:
        # get consecutive occurences of "….."
        bottom_index = indices[count]
        top_index = indices[count + 1]

        if top_index - bottom_index == 1:
            try:
                bottom_index = indices[count + 2]
                top_index = indices[count + 3]
                count += 2
            except:
                break

        # use indices to get row info from the cleaned_list
        row_info = info[bottom_index + 1:top_index]
        clean_row = clean_rows_email(row_info)

        # add row_info to final_list
        final_list.extend(clean_row)
        count += 1

    return (final_list)

# Clean the row info
def clean_rows_email(row_info):
    clean_row= []
    
    # Remove all the elements between the title and the company name
    if "B" in row_info:
        row_info.remove("B")

    if "D" in row_info:
        row_info.remove("D")

    if "HQ" in row_info:
        row_info.remove("HQ")

    if "-" in row_info:
        row_info.remove("-")

    if "-" in row_info:
        row_info.remove("-")

    # Add in the first three elements (name, title, company name)
    first_three = row_info[:3]
    clean_row.extend(first_three)

    # Remove the country
    if ", " in row_info[3]:
        row_info.remove(row_info[3])

    # Check for industry; put "PLACEHOLDER!" if there is not one
    if "43" in row_info[3]:
        clean_row.append("")
        clean_row = []
        return(clean_row)
    else:
        clean_row.append(row_info[3])

    # Find the email and add it to clean_row
    email_indices = [i for i, s in enumerate(row_info) if '@' in s]
    if len(email_indices) > 0:
        email_index = email_indices[0]
        email = row_info[email_index]
        email = email.split(' ', 1)[0]
        clean_row.append(email)
    else:
        clean_row.append("EMAIL NOT FOUND")

    # Find the phone number and add it to clean row (direct if possible, HQ if not)
    phone_indices = [i for i, s in enumerate(row_info) if '(Direct)' in s]
    if len(phone_indices) == 0:
        phone_indices = [i for i, s in enumerate(row_info) if '(HQ)' in s]
    
    if len(phone_indices) > 0:
        phone_number_index = phone_indices[0]
        phone_number = row_info[phone_number_index]
        phone_number = phone_number[:14]
        clean_row.append(phone_number)
    else:
        clean_row.append("PHONE NUMBER NOT FOUND")

    return(clean_row)

# remove rows that are missing information
def no_info_remover(df):
    df_clean = df.dropna()

    return(df_clean)

# Write final info into excel sheet
def organize_info_email():
    paths = get_paths()
    info_list = info_data_email(paths[0])
    final_list = get_rows_email(info_list)

    # elements are in order: name, title, company, industry, email, phone
    # get evert sixth element for each type of info
    name = final_list[0::6]
    title = final_list[1::6]
    company = final_list[2::6]
    industry = final_list[3::6]
    email = final_list[4::6]
    phone = final_list[5::6]

    final_df = pd.DataFrame(list(zip(name, title, company, industry, email, phone)), columns = ["Full Name", "Title", "Company", "Industry", "Email", "Phone Number"])
    final_df = no_info_remover(final_df)
    final_df.to_excel(paths[1], sheet_name='Sheet1', index=False)

    
organize_info_email()
