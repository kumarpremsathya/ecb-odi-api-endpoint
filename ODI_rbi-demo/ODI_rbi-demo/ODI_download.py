from selenium.webdriver.common.by import By
import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import requests

chrome_driver_path = 'C:\Chrome\chrome-win64'
url = 'https://rbi.org.in/Scripts/Data_Overseas_Investment.aspx'

chrome_options = Options()
chrome_options.add_argument("--start-maximized")

download_folder = os.path.join(os.getcwd(), 'years')

if not os.path.exists(download_folder):
    os.makedirs(download_folder)

prefs = {
    "download.default_directory": download_folder,
    "directory_upgrade": True
}

chrome_options.add_experimental_option('prefs', prefs)

driver = webdriver.Chrome(options=chrome_options)
driver.get(url)


start_year = 2012 

archive = driver.find_element(By.XPATH, "//div[@id='divArchiveMain']")
if archive:
    archive.click()
    time.sleep(3) 

years = driver.find_elements(By.XPATH, "//a[@class='year']")[::-1]  # Reversed list of years

try:
    start_index = 0
    for year_index in range(len(years)):
        years = driver.find_elements(By.XPATH, "//a[@class='year']")[::-1]
        year_element = years[year_index]
        year_text = year_element.text.strip()

        
        if year_text.isdigit() and int(year_text) >= start_year:
            start_index = year_index
            break 

    for year_index in range(start_index, len(years)):
        years = driver.find_elements(By.XPATH, "//a[@class='year']")[::-1]
        year_element = years[year_index]
        year_text = year_element.text.strip()
        year_element.click()
        time.sleep(2)

        year_folder = os.path.join(download_folder, year_text)
        if not os.path.exists(year_folder):
            os.makedirs(year_folder)

        months = driver.find_elements(By.XPATH, "//a[@class='link2']")

        for month_index in range(len(months)):
            months = driver.find_elements(By.XPATH, "//a[@class='link2']")
            month_element = months[month_index]
            month_text = month_element.text.strip()
            month_element.click()
            time.sleep(2)

            month_folder = os.path.join(year_folder, month_text)
            if not os.path.exists(month_folder):
                os.makedirs(month_folder)

            if len(driver.window_handles) > 1:
                driver.switch_to.window(driver.window_handles[1])

            
            files = driver.find_elements(By.XPATH, "//a[contains(@href, '.xls') or contains(@href, '.pdf')]")
            
            for file in files:
                file_url = file.get_attribute("href")
                file_name = file_url.split("/")[-1]  

               
                if file_url.endswith(('.xls', '.pdf')):
                    file_path = os.path.join(month_folder, file_name)
                    
                    
                    response = requests.get(file_url)
                    with open(file_path, 'wb') as f:
                        f.write(response.content)

            
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

            driver.back()  
            time.sleep(3)
           
except Exception as e:
    print("Type of Error:", e)
finally:
    driver.quit()









# skip_rows = 0
# while True:
#     df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=skip_rows)
#     if df.columns[0] == 'Sr.':
#         break
#     skip_rows += 1

# # Use the row with 'Sr.' as the header
# header_row = skip_rows
# df = pd.read_excel(file_path, sheet_name=sheet_name)
