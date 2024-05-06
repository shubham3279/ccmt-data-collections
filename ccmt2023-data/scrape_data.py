# imports
import pandas
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# data_extraction

# Initialize Chrome WebDriver
driver = webdriver.Chrome()

# Navigate to the webpage
driver.get('https://admissions.nic.in/ccmt/applicant/report/ORCRReport.aspx?boardid=105012321')


# Find the table element with ID 'ORCRGridView'
table = driver.find_element(By.ID, 'ORCRGridView')

# Find all table header elements within the table
th_elements = table.find_elements(By.TAG_NAME, 'th')

# Extract attributes from the table headers
attributes = [th.text for th in th_elements]



# Extract records
records = []

# Wait for the table to load
table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ORCRGridView')))

# Iterate over page numbers
for i in range(1, 17):
    # Find all table rows excluding rows with class 'gridpager'
    rows = table.find_elements(By.TAG_NAME, 'tr')
    
    # Iterate over rows and extract data
    for row in rows:

        # ignoring irrelevant rows and appending row datas to record list
        if row.get_attribute('class') != 'bg-primary' and row.get_attribute('class') != 'gridpager' and 'javascript:__doPostBack' not in row.get_attribute('innerHTML'):
            # Find all table data cells in this row
            cells = row.find_elements(By.TAG_NAME, 'td')
            # Extract text from each cell and append to the record list
            record = [cell.text for cell in cells]
            records.append(record)
    
    # Click on the next page link
    if i != 10:
        try:
            element = driver.find_element(By.LINK_TEXT, str(i+1))
            element.click()
            # Wait for the table to load again
            table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ORCRGridView')))
        except BaseException:
            break  # Break the loop if the next page link is not found
    elif i == 10:
        try:
            element = driver.find_element(By.LINK_TEXT, '...')
            element.click()
            # Wait for the table to load again
            table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ORCRGridView')))
        except BaseException:
            break  # Break the loop if the next page link is not found


# Close the browser
driver.quit()

# creating pandas dataframe

data = pandas.DataFrame(
    records, 
    columns=attributes
)

# exporting data to csv and exxcel format

data.to_csv('ccmt2023.csv', index=False)
data.to_excel('ccmt2023.xlsx', index=False)