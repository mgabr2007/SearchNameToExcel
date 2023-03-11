# SearchNameToExcel
This Python script allows users to search for files with a specific name in a chosen folder and add links to those files in an Excel file. The program features a graphical user interface using the PySimpleGUI package, and error checking to ensure all input fields are filled before starting the search and link creation process. The script has been modified to function as a Streamlit app, which provides a web-based interface for inputting search parameters. The app displays a success message if the links are added to the Excel file successfully.

To use the script, users must have Python 3 and the required packages (xlwt, xlrd, xlutils, and streamlit) installed. After installation, users can run the script by navigating to the folder containing the script and running streamlit run script_name.py in the terminal or command prompt.

The program first prompts the user to select the folder to search in, followed by an input field for the file name to be searched for, and an input field for the Excel file where the links will be added. Once the user provides the necessary information, they can click the "Search and Add Links" button to execute the program.

If files are found with the specified name, the program adds links to those files in the specified Excel file. In case no files are found with the specified name, the program displays a warning message. The use of a graphical user interface and error checking makes the program simple to use and reduces the risk of errors.

In conclusion, this Python script is a useful tool for searching for files and adding links to them in an Excel file. The use of a Streamlit app and graphical user interface makes the program intuitive to use and minimizes the chance of errors.
