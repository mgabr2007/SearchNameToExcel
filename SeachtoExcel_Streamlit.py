import os
import glob
import xlsxwriter
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
        workbook = xlsxwriter.Workbook(file_path)
        worksheet = workbook.add_worksheet()

        # Add search word and links to Excel sheet
        row = 0
        for path in file_paths:
            link = f'=HYPERLINK("{path}")'
            worksheet.write(row, 0, search_name)
            worksheet.write_url(row, 1, path, string=link)
            row += 1

        # Save changes to Excel sheet
        workbook.close()

        # Display success message
        st.success("File links added successfully!")
    else:
        st.warning(f"No files found with name '{search_name}'")

# Create Streamlit app
st.title("Search and Add Links to Excel Sheet")

# Create input fields for folder, search name, and Excel file
folder_path = st.file_uploader("Folder to search in:", type="directory")
search_name = st.text_input("File name to search for:")
file_path = st.text_input("Excel file to write links to:")

# Add button to start search and link creation
if st.button("Search and Add Links"):
    if folder_path and search_name and file_path:
        search_and_add_links(folder_path, search_name, file_path)
    else:
        st.warning("Please fill in all input fields")
