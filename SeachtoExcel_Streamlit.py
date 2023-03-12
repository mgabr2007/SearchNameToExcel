import os
import xlsxwriter

# Get user input for search name and folder path
search_name = input("Enter file name to search for: ")
folder_path = input("Enter folder path to search in: ")
output_file_path = input("Enter output file path: ")

# Create a new Excel workbook and worksheet
workbook = xlsxwriter.Workbook(output_file_path)
worksheet = workbook.add_worksheet()

# Set up column headers in the worksheet
worksheet.write(0, 0, "Search Name")
worksheet.write(0, 1, "File Path")

# Initialize row counter
row = 1

# Search for files in the folder path that match the search name using os.scandir()
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
print("Search complete. Results saved to {}".format(output_file_path))
