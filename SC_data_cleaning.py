import pandas as pd
import re

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
