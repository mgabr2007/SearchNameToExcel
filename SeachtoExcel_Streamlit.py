# Install required packages
# Run this command in your terminal or command prompt:
# pip install xlwt xlrd xlutils streamlit

import os
import glob
import xlwt
import xlrd
from xlutils.copy import copy
import streamlit as st

def search_and_add_links(folder_path, search_name, file_path):
    # Search for file in selected folder
    file_paths = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if search_name in file:
                file_paths.append(os.path.join(root, file))

    # Create link to file and add it to Excel sheet
    if file_paths:
        # Open Excel file
        workbook = xlrd.open_workbook(file_path)
        worksheet = workbook.sheet_by_index(0)

        # Create a copy of the workbook to write to
        new_workbook = copy(workbook)
        new_worksheet = new_workbook.get_sheet(0)

        # Get next empty row in Excel sheet
        next_row = worksheet.nrows

        # Add search word and links to Excel sheet
        for path in file_paths:
            link = f'=HYPERLINK("{path}")'
            new_worksheet.write(next_row, 0, search_name)
            new_worksheet.write(next_row, 1, link)
            next_row += 1

        # Save changes to Excel sheet
        new_workbook.save(file_path)

        # Display success message
        st.success("File links added successfully!")
    else:
        st.warning(f"No files found with name '{search_name}'")

# Create Streamlit app
st.title("Search and Add Links to Excel Sheet")

# Create input fields for folder, search name, and Excel file
folder_path = st.text_input("Folder to search in:")
search_name = st.text_input("File name to search for:")
file_path = st.text_input("Excel file to write links to:")

# Add button to start search and link creation
if st.button("Search and Add Links"):
    if folder_path and search_name and file_path:
        search_and_add_links(folder_path, search_name, file_path)
    else:
        st.warning("Please fill in all input fields")
