import format_data # Extracts and formats the new DPP Data
import merge_data # Merges new DPP workorders to the master sheet

# The main function will call the format_data and merge_data functions
# This will format the new DPP workorders and append them to the master sheet
# A temporary copy of the master sheet is created in the merge function
    # It will be deleted if the merge is successful
def main(master_sheet_filepath, raw_dpp_filepath):
    # Extract and format the new DPP data
    formatted_dpp_filepath = format_data.format_workorders(raw_dpp_filepath)

    # Merge the formatted DPP data into the master sheet
    merge_data.merge(master_sheet_filepath, formatted_dpp_filepath)


