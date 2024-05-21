# Browser-Automation
There are three different scripts for three different websites but the working is same, also three excel sheets of data required by the script provided.

Browser Automation Scripts 
This code is a Python application designed to automate the extraction of billing information from a specific web page using Selenium WebDriver and Tkinter for a graphical user interface (GUI). Here’s a breakdown of its functionality and workflow:

### Key Components and Functions:

1. **Import Statements**:
    - Import necessary libraries: `time`, `pandas`, `tkinter`, `filedialog` from `tkinter`, and Selenium-related modules.
    - `webdriver_manager.chrome` is used to manage the ChromeDriver for Selenium.

2. **`initialize_driver` Function**:
    - Initializes the Chrome WebDriver using the `webdriver_manager`.

3. **`extract_data` Function**:
    - Extracts billing data for a given K-Number.
    - Checks for a "No dues pending" message.
    - Retrieves the "Bill Due Date" and "Amount Payable" from the web page if no dues are found.

4. **`browse_excel_file` Function**:
    - Opens a file dialog to select an Excel file containing K-Numbers.

5. **`process_data` Function**:
    - Main function that orchestrates the data extraction process.
    - Uses `browse_excel_file` to select the Excel file.
    - Initializes the web driver and opens the specified URL.
    - Reads the K-Numbers from the Excel file.
    - Iterates over each K-Number:
        - Fills out the form on the webpage.
        - Submits the form and extracts billing information.
        - Stores the extracted data.
    - Saves the extracted data to a new Excel file.

6. **`main` Function**:
    - Sets up the Tkinter GUI.
    - Creates a button that triggers the `process_data` function.
    - Runs the Tkinter event loop.

### How It Works:

1. **User Interface**:
    - The application starts with a Tkinter window with a button labeled "Process Data".
    - When the button is clicked, it calls `process_data`.

2. **Data Extraction**:
    - The user is prompted to select an Excel file containing K-Numbers.
    - For each K-Number in the file:
        - The application navigates to a specific web page.
        - Enters the K-Number and a placeholder email into a form.
        - Submits the form and waits for the response.
        - Extracts the billing details or identifies if there are no dues pending.
        - Goes back to the form page to process the next K-Number.
    - Extracted data is saved to a new Excel file named "extracted_data.xlsx".

3. **Error Handling**:
    - The code includes try-except blocks to handle potential errors during web interaction and data extraction.

### Summary:

This script automates the process of extracting billing information for multiple K-Numbers by interacting with a web form, collecting the required data, and saving it to an Excel file. The user interface, built with Tkinter, makes it easy to use by providing a simple button to start the process and a file dialog to select the input file.
