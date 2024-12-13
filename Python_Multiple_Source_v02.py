# -*- coding: utf-8 -*-
"""
Created on Tue Dec 10 17:49:40 2024

@author: Franz Guzman
"""

import pandas as pd
import mhi.pscad
import time

# Path to access the Excel file
excel_path = "E:/OneDrive/1. ESTUDIOS - RODRIGO/1. UNICAMP - M. INGENIERÍA ELÉCTRICA/0. Cursos y Capacitaciones/0. Curso_RTDS_2024_UNICAMP/1. Aula 01/Check_List_RTDS_3.xlsx"  # Replace with the path to your file

# Read a specific sheet from the Excel file
one_sheet = pd.read_excel(excel_path, sheet_name='GERADORES') 
print(one_sheet)

# Start and end points for the table data
init_row, init_col = (7, 13)  # Starting point of the data table
end_row, end_col = (one_sheet.shape[0], one_sheet.shape[1]-3)  # End point of the data table
len_row = end_row - init_row
len_col = end_col - init_col
print("Dimension of one_sheet:", one_sheet.shape)  # Dimensions of rows and columns in the sheet
print("Starting point: ", init_row, " - ", init_col, "; Ending point: ", end_row, " - ", end_col)

# Read only specific columns from the sheet (columns 13 to 34)
input_cols = [i for i in range(init_col, end_col)]  # Range from column 13 to 34
print(input_cols)

# Load the data file again using the selected columns
df = pd.read_excel(excel_path, sheet_name='GERADORES', header=init_row, usecols=input_cols)  # Load data to read specific columns
print("Column names: ", df.columns)
print("New dimension of df: ", df.shape)

# =============================================================================
# create a list of column name's from SRC, necessary for load data in PSCAD
col_name_SRC = [
    "Name", "Type", "ZSeq", "Imp", "PS", "R1s", "R1p", "L1p", "R0s",
    "L0s", "MVA", "Vm", "F", "Es", "F0", "Ph", "Iabc", "P", "Q"
]

# =============================================================================

with mhi.pscad.application() as pscad:

    # Load Workspace_Python workspace. Note: charge an especific worskpace .pswx
    worskpace_python = pscad.load("E:\\OneDrive\\1. ESTUDIOS - RODRIGO\\1. UNICAMP - M. INGENIERÍA ELÉCTRICA\\0. Cursos y Capacitaciones\\1. Curso_PSCAD 2024_UNICAMP\\9. MeuPSCAD\\2. Python\\Python_Projects\\Workspace_Python.pswx")
    
    # Load Teste_01 case estudy. Note: charge an especific case study project .pscx
    pscad.load("E:\\OneDrive\\1. ESTUDIOS - RODRIGO\\1. UNICAMP - M. INGENIERÍA ELÉCTRICA\\0. Cursos y Capacitaciones\\1. Curso_PSCAD 2024_UNICAMP\\9. MeuPSCAD\\2. Python\\Python_Projects\\Sources.pscx")

    # Go to 'Sources' case study
    Sources = pscad.project("Sources") # Go to case study 'Sources'
    canvas_Sources = Sources.canvas("Main") # Go to canvas case study 'Sources'
    Sources.navigate_to() # Navigate to 'Sources'
  
    # Reference coordinates of the element
    elements_in_line = 0  # Counter for the number of elements (sources) in a row
    x0_position, y0_position, theta = 7,7, -45 # referene x0, y0 position
    delta_x, delta_y = 6, 10  # Spacing between elements (sources)
    x_position, y_position = x0_position, y0_position  # Initial position to create a Source in RSCAD

    # depure headers
    col_name_load_str = [str(header).strip() for header in df.columns] # Convert to string and remove spaces
    col_name_SRC_str = [str(header).strip() for header in col_name_SRC] # Convert to string and remove spaces
    
    # Verify the values header from excel is equal to input data in PSCAD 
    state_data_load = True
    
    for index, (header_SRC, header_load) in enumerate(zip(col_name_SRC_str, col_name_load_str)): # Create a dictionary with col_name_SRC_str and col_name_load_str using zip 
        if header_SRC != header_load:
            print(f"__________________Mismatch at index {index}___________________")
            print(f"The column name '{header_load}' is different to '{header_SRC}' in excel")
            print(f"Please change the column name in Excel from: '{header_load}' to -> '{header_SRC}'")
            print("_________________________________________________________")
            state_data_load = False
    
# =============================================================================
    keys_to_round = ["R1s", "R1p", "L1p", "R0s", "L0s", "Es", "Ph"] # Parameters that need around decimals
    around_digits_number = 6 # '6' digist to around
# =============================================================================
    
    if state_data_load == True:
        for data in range(df.shape[0]):
            data_row = df.iloc[data]  # Retrieve data for each row during the iteration
            print(f"Source {data}:")

            # Depure values from excel before to export to PSCAD
            # value_str = [str(header).strip() for header in data_row] # Convert to string and remove spaces
            value_str = [
            str(round(header, around_digits_number)) if header_key in keys_to_round and isinstance(header, (float, int)) else str(header).strip() for header_key, header in zip(col_name_SRC_str, data_row)
            ] # Convert to string and remove spaces, and round to '6' decimal data for keys_to_round

            # Create dictionary
            parameters_dict = dict(zip(col_name_SRC_str, value_str)) # Create other dictionary with col_name_SRC_str and values_str that are ready
            # print("dictionary_created: ", parameters_dict)

            try:
                # Create a SOURCE 'master:source3' is differente to 'master:source_3'. Note: the name from element appear below to instance
                master_source3 = canvas_Sources.create_component("master:source3", x=x_position, y=y_position, orient=theta) 
                master_source3_str= str(master_source3)  # Get de string value
                SRC_id = int(master_source3_str.split("#")[1]) # Identify the element's by unique ID
                master_source3_SRC_id = Sources.component(SRC_id)  # Select the component by id
            except Exception as e: # If not is a source created returs an exception and break
                state_data_load == False
                print(f"Error creating a Source {data}: {e}")
                print("------------The process was interrupted------------")
                break
                
            # Adjust spacing between elements
            x_position = x_position + delta_x 
            elements_in_line += 1 
            
            if elements_in_line == 8:  # Ensure a maximum of 8 elements per row
                y_position = y_position + delta_y # Ensure the height elements per row
                elements_in_line = 0 # Reset elements in line
                x_position = x0_position # Reset position 'x'
            
            for key, value in parameters_dict.items(): 
                try:
                    master_source3_SRC_id.parameters(**{key: value}) # create a dinamic code por execute in python for set paramenters 
                    print(f"{key:<10}: {value:<20} -> ok")    
                    
                except Exception as e:
                    state_data_load = False
                    print(f"{key:<10}: {value:<20} -> Error: {e}")
                    print("------------The process was interrupted------------")
                    break
                
            if not state_data_load:
                break  # if state_data_load = False, break de if.   
                
            
    print("...Simulation_Finished_Successfully...")
    