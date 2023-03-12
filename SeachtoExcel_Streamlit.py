import os
import re
import streamlit as st
import pandas as pd


def search_files(search_name, folder_path, output_file_path):
    # Create an empty DataFrame to store the search results
    results_df = pd.DataFrame(columns=["Search Name", "File Path"])

    # Set up search criteria
    search_criteria = ".*{}.*".format(search_name)

    # Search for files in the folder path that match the search name using os.scandir()
    with os.scandir(folder_path) as entries:
        for entry in entries:
            if entry.is_file() and re.search(search_criteria, entry.name):
                # Add the search name and file path to the results DataFrame
                results_df = results_df.append(
                    {"Search Name": search_name, "File Path": entry.path}, ignore_index=True
                )

    # Write the results to the output Excel file
    results_df.to_excel(output_file_path, index=False)

    # Print a message to indicate that the search is complete
    st.success(f"Search complete. Results saved to {output_file_path}")


# Set up the Streamlit app
st.title("Search for Files")
search_name = st.text_input("Enter file name to search for:")
folder_path = st.file_input("Enter folder path to search in:", type="directory")
output_file_path = st.file_uploader("Select output Excel file:", type=["xlsx"])

# Wait for the user to select an output file
if output_file_path is not None:
    # Run the search function with the user inputs
    search_files(search_name, folder_path, output_file_path.name)
