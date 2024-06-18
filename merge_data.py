from openpyxl import *
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
from datetime import date
import os

# Merge formatted DPP data into the Master Excel sheet
def merge(master_sheet_filepath, formatted_dpp_filepath):
    
    # Flag to check for successful merge
    # This will determine if the master sheet copy will be removed or not
    success = False
    
    # Make a copy of the master sheet for backup
    master_copy_filepath = master_sheet_filepath
    sheet_name = "Sheet0"
    df = pd.read_excel(master_sheet_filepath, sheet_name=sheet_name)
    with pd.ExcelWriter(master_copy_filepath) as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Open the master sheet
    master_df = load_workbook(master_sheet_filepath)
    master_ws = master_df[sheet_name]
    
    # Extract the existing dispatch numbers from the master sheet
    master_dispatch_numbers = df["Dispatch Number"].tolist()

    # Open the DPP sheet
    dpp_df = load_workbook(formatted_dpp_filepath)
    dpp_ws = dpp_df["Sheet0"]

    # Iterate through the rows in the DPP data and extract new dispatch numbers (work orders)
    new_workorders = []
    for index, row in dpp_ws.iterrows():
        if row["Dispatch Number"] not in master_dispatch_numbers:
            new_workorders.append(row)
    
    # Create a new DataFrame for the new work orders
    new_workorders_df = pd.DataFrame(new_workorders)
    new_workorders_df = new_workorders_df.active # Assumes data is in the active sheet
   
    # Append the new work orders to the master sheet
    # Iterate through new_workorders_df and append the new data (skip headers)
    for row in dataframe_to_rows(new_workorders_df, index=False, header=False):
        master_ws.append(row) 

    # Save the updated master sheet
    master_ws.save(master_sheet_filepath)

    # Set the success flag to True
    # This will remove the master sheet copy
    print('The DDP data was successfully merged.')
    success = True

    # If successful, remove the master sheet copy
    if success:
        try:
            os.remove(master_copy_filepath)
            print('Master sheet copy removed successfully.')
        except FileNotFoundError:
            print('Master sheet copy not found.')
        except PermissionError:
            print('Master sheet copy is open.')
        except Exception as e:
            print(f'An unexpected error occurred: {e}.')
