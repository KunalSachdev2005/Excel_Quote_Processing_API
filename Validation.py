#******************************************************************************************
# Program Name: Validation.py
# Description: This progarm validates the quote excel file uploaded by the user to the API.
# Created on 01-06-2022 by Kunal Sachdev
#******************************************************************************************


def validation(file_path):
    """This function validates the excel file given by the user"""
    
    # Importing required libraries
    import pandas as pd
    import os
    from openpyxl import load_workbook
    
    # Reading Excel file in Python
    oricon_df = pd.read_excel(sheet_name = "Quotation Sheet", io = file_path, nrows = 100, usecols = list(range(0, 51)), header = None)
    
    # Reading properties of excel file
    wb = load_workbook(file_path)
    prop = wb.properties
    
    # Encoded Value
    encoded_value = '2f9a0eddcb63e177690adf5c931acecb'
    
    #flags
    flag_cells = False
    flag_encoded = False
    flag_prop = False
    
    # Validating Some Cell Values and encoded value
    if oricon_df[0][5] == 'Hazard Category' and oricon_df[0][7] == 'NAT CAT Score' and oricon_df[0][8] == 'Graded Retention' and oricon_df[0][26] == 'Mailing Address: ' and oricon_df[0][28] == 'Occupancy:':
        flag_cells = True
    
    # Validating encoded value
    if oricon_df[14][0] == encoded_value:
        flag_encoded = True
    
    # Validating file properties
    if prop.title == 'BSC' and prop.creator == 'Rakesh Kedia':
        flag_prop = True
            
    # Checking if flags are true        
    if flag_cells == True and flag_encoded == True and flag_prop == True:
        return True
    else:
        return False