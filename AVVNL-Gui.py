import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def initialize_driver():
    # Initialize Chrome WebDriver
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service)

def extract_data(driver, knumber):
    try:
        # Check if "No dues pending" message is present
        no_dues_element = driver.find_elements(By.XPATH, "/html/body/form/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]/table/tbody[1]/tr[1]/td/table/tbody/tr[3]/td[2]")
        if no_dues_element:
            return {"K-Number": knumber, "Bill Due Date": "No dues pending", "Amount Payable": "No dues pending"}
        else:
            # Wait for the elements to be present
            bill_due_date = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form/div/div[2]/table/tbody/tr[9]/td[2]')))
            amount_payable = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form/div/div[2]/table/tbody/tr[11]/td[2]')))
            # Extract Bill Due Date and Amount Payable
            return {"K-Number": knumber, "Bill Due Date": bill_due_date.text, "Amount Payable": amount_payable.text}
    except Exception as e:
        print(f"An error occurred while processing K number {knumber}: {str(e)}")
        return {"K-Number": knumber, "Bill Due Date": "Error", "Amount Payable": "Error"}

def browse_excel_file():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    return file_path

def process_data():
    # Browse for Excel file
    excel_file = browse_excel_file()
    if not excel_file:
        print("No file selected. Exiting...")
        return

    # Open the given link
    link = "https://pgi.billdesk.com/pgidsk/pgmerc/rvvnlaj/RVVNLAJDetails.jsp"
    driver = initialize_driver()
    driver.get(link)
    time.sleep(0.2)  # Adjust as needed for the page to load

    # Read K numbers from Excel file
    df = pd.read_excel(excel_file)
    knumbers = df["K-NUMBER"].tolist()

    # Iterate through K numbers and process each one
    extracted_data = []
    for kno in knumbers:
        try:
            # Fill K number and email fields
            k_number_field = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form/div/div[3]/table/tbody/tr[4]/td[2]/input')))
            k_number_field.clear()
            k_number_field.send_keys(str(kno))

            email_field = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form/div/div[3]/table/tbody/tr[5]/td[2]/input[1]')))
            email_field.clear()
            email_field.send_keys("xyz@gmail.com")

            # Submit the form
            submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div/div[3]/table/tbody/tr[5]/td[2]/input[2]')))
            submit_button.click()

            # Wait for the next page to load and extract data
            time.sleep(0.2)  # Adjust as needed
            extracted_data.append(extract_data(driver, kno))

            # Go back to the landing webpage
            driver.back()
            time.sleep(0.2)  # Adjust as needed for the page to load again
        except Exception as e:
            print(f"An error occurred while processing K number {kno}: {str(e)}")
    
    # Close the browser
    driver.quit()

    # Save extracted data to a new Excel file
    extracted_df = pd.DataFrame(extracted_data)
    extracted_df.to_excel("extracted_data.xlsx", index=False)

def main():
    # Create the main window
    root = tk.Tk()
    root.title("Data Extraction Tool- Ajmer-AVVNL")

    # Set the size of the window
    root.geometry("400x200")  # Set width x height according to your preference

    # Create a button to trigger the data extraction process
    process_button = tk.Button(root, text="Process Data", command=process_data, width=20 ,height=10)
    process_button.pack(side=tk.TOP, padx=10, pady=10)

    # Run the Tkinter event loop
    root.mainloop()

if __name__ == "__main__":
    main()
