import os
import xlsxwriter
import streamlit as st

# Get user input for search name and folder path
search_name = st.text_input("Enter file name to search for:")
folder_path = st.text_input("Enter folder path to search in:")
output_file_path = st.text_input("Enter output file path:")

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
if folder_path:
    with os.scandir(folder_path) as entries:
        for entry in entries:
            if entry.is_file() and search_name in entry.name:
                # Write the search name and file path to the worksheet
                worksheet.write(row, 0, search_name)
                worksheet.write(row, 1, entry.path)
                row += 1

    # Close the Excel workbook
    workbook.close()

    # Print a message to indicate that the search is complete
    st.write("Search complete. Results saved to {}. ".format(output_file_path))

    # Display the Excel file as a DataFrame in Streamlit
    df = pd.read_excel(output_file_path)
    st.write(df)
else:
    st.write("Please enter a valid folder path to search in.")
