import os
import xlsxwriter
import streamlit as st

# Set up the Streamlit app
st.title("File Search to Excel")
st.write("Enter search criteria and select folder to search in.")

# Get user input for search name and folder path
search_name = st.text_input("Enter file name to search for:")
folder_path = st.file_input("Enter folder path to search in:", type="directory")
output_file_path = st.file_uploader("Select output file path:", type=["xls", "xlsx"])

# If the user has selected a valid output file path, create a new Excel workbook and worksheet
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
    st.write("Search complete. Results saved to {}.".format(output_file_path))
