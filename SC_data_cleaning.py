import pandas as pd
import re
import numpy as np

# /home/shane/Downloads/SHEPHERDS CLINIC 083120.xlsx

raw_path = input("Give the absolute filepath of the sheet to be converted:") # get pathname (note MUST BE .XLSX file)

raw_data = pd.read_excel(io = raw_path) # Read in the .xlsx file

row_number = raw_data.shape[0] #get the number of rows in the dataframe

new_df = pd.DataFrame(columns = ['full_name', "drug"]) # Make a new dataframe to amend results into

for i in range(row_number): #Iterate through every row
    cell_string = str(raw_data["Unnamed: 0"][i]) # Make the cell you're searching from column 1 a string to be searched
    if bool(re.search(", ", cell_string)) & (not bool(re.search("Subtotal", cell_string))): # If the cell of interest contains a comma but does NOT contain "Subtotal", this means that we have identified the entry of a patient

        name = str(raw_data["Unnamed: 0"][i])
        name_list = []
        drug_list = []
        counter = 1
        drug_entry = str(raw_data["Unnamed: 4"][i + counter])

        while drug_entry != 'nan':
            name_list.append(name)
            drug_list.append(drug_entry)
            counter += 1
            drug_entry = str(raw_data["Unnamed: 4"][i + counter])

        new_entries = pd.DataFrame({
            "full_name": name_list,
            "drug": drug_list
        })

        new_df= pd.concat([new_df, new_entries]).reset_index(drop = True)

    else: #if this
        pass

name_split_df = new_df['full_name'].str.split(pat = ", ", expand = True)

final_df = pd.concat([new_df.drop(labels = "full_name", axis = 1), name_split_df], axis = 1)

final_df = final_df.rename(columns = {0: "Last", 1: "First"})

final_df["Last"] = final_df["Last"].str.strip()

final_df["First"] = final_df["First"].str.strip()

final_df["Date_filled"] = final_df['drug'].str.extract("([0-9]{2}/[0-9]{2}/[0-9]{4})")

final_df['Medication'] = final_df['drug'].str.extract("((?<=\ - )(.*?)(?=\\(DateFilled))", expand = False)[0]

final_df['Dosage'] = final_df['Medication'].str.extract("([0-9].*$)")

final_df['Medication'] = final_df['Medication'].str.extract("(^.*?(?=[0-9]))")

final_df['Grant'] = np.NaN

final_df['Pharmacy'] = np.NaN

final_df['Refill_Date'] = np.NaN

final_df['Notes'] = np.NaN

final_df['Include'] = "Y"

final_df = final_df.drop(columns = ['drug'])

final_file_path = input("Give the absolute filepath of the desired output file (example: C:\Documents\excel_sheets\August_data_cleaned.xlsx):")

final_df.to_excel(final_file_path, index = False)
