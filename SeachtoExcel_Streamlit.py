import os
import xlsxwriter
import time
import streamlit as st

# Get user input for search name and folder path
search_name = st.text_input("Enter file name to search for:")
output_file_path = st.file_uploader("Select output Excel file:", type=["xlsx", "xls"])

# Create a new Excel workbook and worksheet
if output_file_path is not None:
    workbook = xlsxwriter.Workbook(output_file_path)
    worksheet = workbook.add_worksheet()

    # Set up column headers in the worksheet
    worksheet.write(0, 0, "Search Name")
    worksheet.write(0, 1, "File Path")

    # Initialize row counter
    row = 1

    # Set up search criteria
    search_criteria = ".*{}.*".format(search_name)

    # Start timer
    start_time = time.time()

    # Search for files in the folder path that match the search name using os.scandir()
    folder_path = st.file_input("Enter folder path to search in:", type="directory")
    if folder_path is not None:
        with os.scandir(folder_path) as entries:
            for entry in entries:
                if entry.is_file() and re.search(search_criteria, entry.name):
                    # Write the search name and file path to the worksheet
                    worksheet.write(row, 0, search_name)
                    worksheet.write(row, 1, entry.path)
                    row += 1

        # Close the Excel workbook
        workbook.close()

        # Calculate elapsed time
        elapsed_time = time.time() - start_time

        # Print a message to indicate that the search is complete
        st.write("Search complete. Results saved to {}. Elapsed time: {:.2f} seconds".format(output_file_path.name, elapsed_time))
