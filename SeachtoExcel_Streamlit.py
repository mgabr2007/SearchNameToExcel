import os
import xlsxwriter
import streamlit as st

# Get user input for search name, folder path, and output file path
search_name = st.text_input("Enter file name to search for:")
folder_path = st.file_input("Enter folder path to search in:", type="directory")
output_file_path = st.file_uploader("Select output Excel file:", type=["xlsx", "xls"])

# Check if user has entered search name and selected a folder and output file
if search_name and folder_path and output_file_path:
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

    # Start searching for files in the folder path that match the search name using os.scandir()
    with os.scandir(folder_path) as entries:
        for entry in entries:
            if entry.is_file() and search_name.lower() in entry.name.lower():
                # Write the search name and file path to the worksheet
                worksheet.write(row, 0, search_name)
                worksheet.write(row, 1, entry.path)
                row += 1

    # Close the Excel workbook
    workbook.close()

    # Show success message to the user
    st.success("Search complete. Results saved to {}".format(output_file_path.name))

else:
    # Show error message to the user if any of the required input is missing
    st.error("Please enter search name, select folder path, and upload an output Excel file.")
