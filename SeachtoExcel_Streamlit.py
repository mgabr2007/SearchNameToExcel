import os
import glob
import xlsxwriter

# Get user input for search name and folder path
search_name = input("Enter file name to search for: ")
folder_path = input("Enter folder path to search in: ")
output_file_path = input("Enter output file path: ")

# Use raw string literal to properly represent the search name
search_name = r"{}".format(search_name)

# Create a new Excel workbook and worksheet
workbook = xlsxwriter.Workbook(output_file_path)
worksheet = workbook.add_worksheet()

# Set up column headers in the worksheet
worksheet.write(0, 0, "Search Name")
worksheet.write(0, 1, "File Path")

# Initialize row counter
row = 1

# Search for files in the folder path that match the search name
for file_path in glob.glob(os.path.join(folder_path, "**", "*{}*".format(search_name)), recursive=True):
    # Write the search name and file path to the worksheet
    worksheet.write(row, 0, search_name)
    worksheet.write(row, 1, file_path)
    row += 1

# Close the Excel workbook
workbook.close()

# Print a message to indicate that the search is complete
print("Search complete. Results saved to {}".format(output_file_path))
