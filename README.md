# dpp_script

## Installation Steps
- Create a virtual environment
    python3 -m venv env OR python -m venv env (depending on your Python version)

- Setup
    1) Navigate to your program directory:     cd dpp_script
    2) Activate Python virtual environment:    source env/bin/activate
    3) Install packages:                       pip install pandas openpyxl datetime

## Files
    - main.py
        - The entry point to the DPP script
        - It calls the format_data and merge_data files to handle data processing

    - format_data.py
        - Takes the raw DPP Excel sheet as an argument
            - Reads through raw file
            - Generates a new Excel sheet with expected formatting and styling
            - returns the filename of the new Excel sheet to be merged with the Master Excel sheet
        
    - merge_data.py
        - Takes the formatted DPP Excel sheet and Master Excel sheet and merges them together
        - It searches through the existing work orders and only adds new work orders (the raw DPP file has many additional entries that will have added to Master previously)