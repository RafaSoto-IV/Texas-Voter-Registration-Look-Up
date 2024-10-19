from selenium import webdriver
from selenium.webdriver.chrome.service import service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import os
import chromedriver_autoinstaller
import openpyxl
from openpyxl.styles import Font
import pandas as pd
from datetime import datetime


def file_selection():
    direct = os.getcwd()
    excel_files = []
    print("Excel Files Found:")
    index = 0
    for filename in os.listdir(direct):
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            excel_files.append(filename)
            print("[" + str(index) + "] " + filename)
            index += 1

    while True:
        try:
            print("Please enter the number associated with the file you would like to use: ")
            file_index = input()
            direct = os.getcwd() + "\\" + str(excel_files[int(file_index)])
        except ValueError:
            print("Invalid input.")
        except IndexError:
            print("Number is not in range.")
        else:
            break

    return direct


def setup(direct):
    # Step 1: Load the Excel file (for reading)
    workbook = openpyxl.load_workbook(direct)
    df = pd.read_excel(direct, usecols=['First Name', 'Last Name', 'DOB', 'County', 'ZIP', 'Status'])

    # Step 2: Collect the relevant columns, storing each row as a tuple inside a list
    # As this iterates it should search and edit/store edits to see if it meets being registered
    data = []
    for index, row in df.iterrows():

        if str(row['First Name']).lower() != 'nat' and str(row['First Name']).lower() != '' and str(row['First Name']).lower() != 'nan':
            dob_variable = str(row['DOB']).replace(" 00:00:00", "")
            dob_string = ''
            while True:
                try:
                    date_object = datetime.strptime(dob_variable, "%Y-%m-%d")
                    dob_string = date_object.strftime("%m/%d/%Y")
                except ValueError:
                    dob_string = dob_variable
                    break
                else:
                    break

            row_data = [
                str(row['First Name']), # First name
                str(row['Last Name']),  # Last name
                dob_string,  # Date of Birth
                str(row['County']),  # County
                str(row['ZIP']), # Zip code
                str(row['Status']) # Status
            ]

        else:
            row_data = [
                'NaT',  # Last name
                'NaT',  # First name
                'NaT',  # Date of Birth
                'NaT',  # County
                'NaT',  # Zip code
                str(row['Status'])   # Status
            ]
        data.append(row_data)
    return data


def iterate(data):
    # Go to Website.
    chromedriver_autoinstaller.install()
    driver = webdriver.Chrome()
    y = 0
    index = 0
    while index < len(data):
        if str(data[index][0]).lower() != 'nat' and str(data[index][0]).lower() != '' and str(data[index][0]).lower() != "nan":
            driver.get('https://teamrv-mvp.sos.texas.gov/MVP/back2HomePage.do')

            first_name = str(data[index][0]).replace("-", "").replace("'", "").replace("’", "")
            last_name = str(data[index][1]).replace("-", "").replace("'", "").replace("’", "")
            dob = data[index][2]
            county = data[index][3]
            zip_code = data[index][4]

            selection_criteria = driver.find_element(By.NAME, 'selType')
            selection_criteria.send_keys('n')
            first_name_box = driver.find_element(By.NAME, 'firstName')
            first_name_box.send_keys(first_name)
            last_name_box = driver.find_element(By.NAME, 'lastName')
            last_name_box.send_keys(last_name)
            county_box = driver.find_element(By.NAME, 'county')
            county_box.send_keys(county)
            dob_box = driver.find_element(By.NAME, 'dob')
            dob_box.send_keys(dob)
            zip_code_box = driver.find_element(By.NAME, 'adZip5')
            zip_code_box.send_keys(zip_code)

            submit_button = driver.find_element(By.ID, 'VALIDBTN')
            submit_button.click()
            try:
                WebDriverWait(driver, 1).until(EC.alert_is_present(),
                                                'Timed out waiting for PA creation ' +
                                                'confirmation popup to appear.')

                alert = driver.switch_to.alert
                alert.accept()
            except TimeoutException:
                print()

            '''print('First Name: ' + str(row[0]))
            print('Last Name: ' + str(row[1]))
            print('DOB: ' + str(row[2]))
            print('County: ' + str(row[3]))
            print('ZIP: ' + str(row[4]))
            print('Status: ' + str(row[5]))'''

            elements = driver.find_elements(By.TAG_NAME, 'span')
            for element in elements:
                if element.text == 'Voter Status: ACTIVE':
                    data[index][5] = 'V'

        index += 1

    return data


def writing (extracted_data, direct):
    workbook = openpyxl.load_workbook(direct)
    sheet = workbook.active
    green_font = Font(bold=True, color="34a853")
    index = 2
    index_offset = 2
    while index <= len(extracted_data) + 1:
        sheet['A' + str(index)] = str(extracted_data[index - index_offset][5])
        if str(extracted_data[index - index_offset][5]) == 'V':
            for cell in sheet[str(index) + ':' + str(index)]:
                cell.font = green_font
        index += 1
    workbook.save(direct)


def main():
    direct = file_selection()
    data = setup(direct)
    extracted_data = iterate(data)
    writing(extracted_data, direct)


main()