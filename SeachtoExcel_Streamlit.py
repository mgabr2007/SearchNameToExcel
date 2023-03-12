import os
import xlsxwriter
import streamlit as st

# Get user input for search name and folder path
search_name = st.text_input("Enter file name to search for: ")
folder_path = st.file_input("Enter folder path to search in: ", type="directory")
output_file_path = st.file_uploader("Enter output file path: ", type=["xlsx"])

# Create a new Excel workbook and worksheet
workbook = xlsxwriter.Workbook(output_file_path)
worksheet = workbook.add_worksheet()

# Set up column headers in the worksheet
worksheet.write(0, 0, "Search Name")
worksheet.write(0, 1, "File Path")

# Initialize row counter
row = 1

# Set up search criteria
search_criteria = ".*{}.*".format(search_name)

# Search for files in the folder path that match the search name using os.scandir()
with os.scandir(folder_path) as entries:
    for entry in entries:
        if entry.is_file() and re.search(search_criteria, entry.name):
            # Write the search name and file path to the worksheet
            worksheet.write(row, 0, search_name)
            worksheet.write(row, 1, entry.path)
            row += 1

# Close the Excel workbook
workbook.close()

# Print a message to indicate that the search is complete
st.success("Search complete. Results saved to {}. Elapsed time: {:.2f} seconds".format(output_file_path, elapsed_time))
