import os
import time
import sys
import traceback
import pandas as pd
import mysql.connector
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException



# Define your database configuration
db_config = {
    'host': 'localhost',
    'user': 'root',
    'password': 'root',
    'database': 'rbi_odi',
    'auth_plugin': 'mysql_native_password'
}

conn = mysql.connector.connect(**db_config)
cursor = conn.cursor()


log_list = [None] * 11


def get_data_count(cursor):
    cursor.execute("SELECT COUNT(*) FROM odi")
    return cursor.fetchone()[0]

total_count_first = get_data_count(cursor)

def get_current_datetime():
    current_datetime = datetime.now()
    return current_datetime.strftime("%Y-%m-%d %H:%M:%S")

current_datetime = get_current_datetime()




source_name = "rbi_odi"
date_of_scraping = current_datetime
log_list[0] = source_name
log_list[10] = date_of_scraping

actual_months = ['November', 'December', 'June', 'September', 'February', 'July', 'August', 'October', 'January', 'April', 'May', 'March']

outer_month_xpath = {"January": "(//tbody//tr//a[contains(text(),'January')]//ancestor::tr//td//a)[1]",
                  "February": "(//tbody//tr//a[contains(text(),'February')]//ancestor::tr//td//a)[1]",
                  "March": "(//tbody//tr//a[contains(text(),'March')]//ancestor::tr//td//a)[1]",
                  "April": "(//tbody//tr//a[contains(text(),'April')]//ancestor::tr//td//a)[1]",
                  "May": "(//tbody//tr//a[contains(text(),'May')]//ancestor::tr//td//a)[1]",
                  "June": "(//tbody//tr//a[contains(text(),'June')]//ancestor::tr//td//a)[1]",
                  "July": "(//tbody//tr//a[contains(text(),'July')]//ancestor::tr//td//a)[1]",
                  "August": "(//tbody//tr//a[contains(text(),'August')]//ancestor::tr//td//a)[1]",
                  "September": "(//tbody//tr//a[contains(text(),'September')]//ancestor::tr//td//a)[1]",
                  "October": "(//tbody//tr//a[contains(text(),'October')]//ancestor::tr//td//a)[1]",
                  "November": "(//tbody//tr//a[contains(text(),'November')]//ancestor::tr//td//a)[1]",
                  "December": "(//tbody//tr//a[contains(text(),'December')]//ancestor::tr//td//a)[1]"}

download_xpath = "//a[contains(@href, '.xls') or contains(@href, '.pdf')]"


def insert_data_into_database(cursor, data_dict,new_file_name):
    global log_list
    # Iterate through rows and append rows with at least min_columns data points to the list

    number_of_rows = len(data_dict)
    log_list[2] = number_of_rows
    #print(data_dict)

    base_name, extension = os.path.splitext(new_file_name)
    month, year = base_name.split()

    for row in data_dict:
        # Check if the length of the row is at least 9
        if len(row) >= 9:
            # Remove 'nan' values from the row and convert it to a dictionary
            row_data = {
                'Name_of_the_Indian_Party': row[1],
                'Name_of_the_JV_WOS': row[2],
                'Whether_JV_WOS': row[3],
                'Overseas_Country': row[4],
                'Major_Activity': row[5],
                'Equity': row[6],
                'Loan': row[7],
                'Guarantee_Issued': row[8],
                'Total': row[9],
                'Month': month,
                'Year': year
            }

            # Insert the data into the database
            query = """
                    INSERT INTO odi (Name_of_the_Indian_Party, Name_of_the_JV_WOS,  Whether_JV_WOS, Overseas_Country, Major_Activity, Equity, Loan, Guarantee_Issued, Total, Month, Year)
                    VALUES (%(Name_of_the_Indian_Party)s, %(Name_of_the_JV_WOS)s, %(Whether_JV_WOS)s, %(Overseas_Country)s, %(Major_Activity)s, %(Equity)s, %(Loan)s, %(Guarantee_Issued)s, %(Total)s, %(Month)s, %(Year)s)
                """
            cursor.execute(query, row_data)
        else:
            print(f"Skipping row with insufficient data: {row}")

    conn.commit()



# def insert_data_into_database(cursor, data_dict, month, year):
#     global log_list
#     # Iterate through rows and append rows with at least min_columns data points to the list


#     number_of_rows = len(data_dict)
#     log_list[2] = number_of_rows
#     print(data_dict)




#     for row in data_dict:
#         # Remove 'nan' values from the row and convert it to a dictionary
#         row_data = {
#             'Name_of_the_Indian_Party': row[1],
#             'Name_of_the_JV_WOS': row[2],
#             'Whether_JV_WOS': row[3],
#             'Overseas_Country': row[4],
#             'Major_Activity': row[5],
#             'Equity': row[6],
#             'Loan': row[7],
#             'Guarantee_Issued': row[8],
#             'Total': row[9],
#             'Month' : month,
#             'Year' : year
#         }

#         # Insert the data into the database
#         query = """
#                 INSERT INTO odi (Name_of_the_Indian_Party, Name_of_the_JV_WOS,  Whether_JV_WOS, Overseas_Country, Major_Activity, Equity, Loan, Guarantee_Issued, Total, Month, Year)
#                 VALUES (%(Name_of_the_Indian_Party)s, %(Name_of_the_JV_WOS)s, %(Whether_JV_WOS)s, %(Overseas_Country)s, %(Major_Activity)s, %(Equity)s, %(Loan)s, %(Guarantee_Issued)s, %(Total)s, %(Month)s, %(Year)s)
#             """
#         cursor.execute(query, row_data)

#     conn.commit()

def get_excel_data_with_min_columns(cursor, excel_file_path, new_file_name,min_columns=8):
    # Read Excel data into a pandas DataFrame, skipping sheets with names 'Summary', 'summary', and 'SUMMARY'
    df = pd.read_excel(excel_file_path, sheet_name=None)
   
    # Get the names of all sheets in the Excel file
    sheet_names = df.keys()

    selected_rows = []

    # Iterate through sheets and rows, and append rows with at least min_columns data points to the list
    for sheet_name in sheet_names:
        # Convert sheet_name to lowercase for case-insensitive comparison
        lower_sheet_name = sheet_name.lower()

        # Skip sheets with names 'summary', 'summary', and 'SUMMARY'
        if lower_sheet_name == 'summary':
            continue

        sheet_df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        for index, row in sheet_df.iterrows():
            if row.count() >= min_columns:
                row_list = row.dropna().tolist()
                selected_rows.append(row_list)

    insert_data_into_database(cursor, selected_rows, new_file_name)






def download_files_for_year(year,outer_month_xpath,download_xpath,cursor,remaining_months):
    global log_list
 
    download_base_directory = fr"C:\Users\mohan.7482\Desktop\EXTERNAL COMMERCIL BORROWING\years"
    download_directory = os.path.join(download_base_directory, str(year))

    # Check if the year folder exists, create it if not
    if not os.path.exists(download_directory):
        os.makedirs(download_directory)

    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_directory,
        "download.default_content_setting_value": 2,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    try:
        browser = webdriver.Chrome(options=chrome_options)

        browser.get("https://rbi.org.in/Scripts/Data_Overseas_Investment.aspx#")
        time.sleep(5)

        archive = browser.find_element(By.XPATH, '//*[@id="divArchiveMain"]')
        archive.click()
        time.sleep(5)

        try:
            next_xpath = browser.find_element(By.XPATH, f'//*[@id="{year}"]')
            next_xpath.click()
            time.sleep(5)
        except Exception as e:
            print(year,"is not updated in the website")
            status_log = "Failure"
            failure_reason = f"{year} is not updated in the website"
            table_name_log = "rbi_odi"
            current_datetime = get_current_datetime()
            date_of_scraping = current_datetime
            total = get_data_count(cursor)
            log_list[0] = table_name_log
            log_list[1] = status_log
            log_list[4] = total
            log_list[8] = failure_reason
            log_list[10] = date_of_scraping
            print(log_list)
            insert_log_into_table(cursor, log_list)
            conn.commit()
            log_list = [None] *11
            # traceback.print_exc()
            sys.exit("Error occurred. Exiting the program.")


        for month in remaining_months:
            total_count = get_data_count(cursor)



            
            next_xpath = browser.find_element(By.XPATH, f'//*[@id="{year}"]')
            next_xpath.click()
            time.sleep(5)
           


            try: 
                time.sleep(5)
                xpath = outer_month_xpath[f"{month}"]
                month_xpath = browser.find_element(By.XPATH,xpath)
                new_file_name = month_xpath.text.split("for")[-1].strip()
                print(new_file_name)
                month_xpath.click()
                time.sleep(5)
            except Exception as e:
                print(month,year,"File is not in website")
                # print("error occured:",e)
                browser.back()
                continue
            




            original_xpath = download_xpath
            original_filename = browser.find_element(By.XPATH,original_xpath).get_attribute("href").split("/")[-1]
            original_filename = original_filename.replace("%20", " ")
            print("original file name as in the website is ", original_filename)
            
         
            original_extension=os.path.splitext(original_filename)[1]
            print("original extension of the file",original_extension)


            new_file_name = new_file_name + original_extension
            print(new_file_name,"new file name want to change")


            downloaded_files = os.listdir(download_directory)
            print(f"the files which are in the {year} folder", downloaded_files)

        
            try:
                if original_filename in downloaded_files or new_file_name in downloaded_files:
                    print(original_filename,"or",new_file_name,"File is already downloaded in that folder")
                    print(original_filename,"or",new_file_name,"File is already inserted into the database")
                    browser.back()
                    continue  
                else:
                    month_download = browser.find_element(By.XPATH, download_xpath)
                    month_download.click()
                    time.sleep(30)
                    old_file_path = os.path.join(download_directory,original_filename)
                    new_file_path = os.path.join(download_directory,new_file_name)
                    os.rename(old_file_path,new_file_path)
                    print(original_filename,"is downloaded and renamed to",new_file_name)
                    time.sleep(15)
                    try: 
                        base_name, extension = os.path.splitext(new_file_name)
                        month_log, year_log = base_name.split()
                        get_excel_data_with_min_columns(cursor, new_file_path, new_file_name)
                        time.sleep(15)
                        print(new_file_name,"is successfully inserted into the datbase")
                        file_name_log = new_file_name
                        table_name_log = "rbi_odi"
                        current_datetime = get_current_datetime()
                        date_of_scraping = current_datetime
                        total = get_data_count(cursor)
                        log_list[0] = table_name_log
                        log_list[4] = total   
                        log_list[5] = month_log
                        log_list[6] = year_log
                        log_list[7] = file_name_log
                        log_list[10] = date_of_scraping
                        incremental_count = get_data_count(cursor)
                        status_log = "Success"
                        log_list[1] = status_log
                        incremental_count = incremental_count - total_count
                        log_list[3] = incremental_count        
                        print(log_list)
                        insert_log_into_table(cursor, log_list)
                        conn.commit()
                        log_list = [None] *11
                        browser.back()
                    except Exception as e:
                        print("error occured:",e)
                        print(year,"is not updated in the website")
                        status_log = "Failure"
                        failure_reason = f"error in insert part for the {new_file_name}"
                        table_name_log = "rbi_odi"
                        current_datetime = get_current_datetime()
                        date_of_scraping = current_datetime
                        total = get_data_count(cursor)
                        log_list[1] = status_log
                        log_list[4] = total 
                        log_list[8] = failure_reason
                        log_list[0] = table_name_log
                        log_list[10] = date_of_scraping
                        print(log_list)
                        insert_log_into_table(cursor, log_list)
                        conn.commit()
                        log_list = [None] *11
                        browser.back()
                        sys.exit(f"error occured in the insert part for this {new_file_name}")

            except Exception as e:
                print("error occured:",e)
                traceback.print_exc()
                browser.back()
                continue
            continue



        

    except TimeoutException:
        print(f"Timeout exception occurred for {year}. Please check the internet connection and try again.")
        table_name_log = "rbi_odi"
        current_datetime = get_current_datetime()
        date_of_scraping = current_datetime
        total = get_data_count(cursor)
        log_list[0] = table_name_log
        log_list[10] = date_of_scraping
        status_log = "Failure"
        reason_log = " 404 error, website is not opened"
        log_list[1] = status_log
        log_list[4] = total 
        log_list[8] = reason_log
        print(log_list)
        insert_log_into_table(cursor, log_list)
        conn.commit()
        log_list = [None] *11
        sys.exit("error occured")
        
        
    except NoSuchElementException:
        print(f"Element not found exception occurred for {year}. Please check if the webpage structure has changed.")
        # table_name_log = "rbi_odi"
        # current_datetime = get_current_datetime()
        # date_of_scraping = current_datetime
        # log_list[0] = table_name_log
        # log_list[9] = date_of_scraping
        # status_log = "Failure"
        # reason_log = " 404 error, website is not opened"
        # log_list[1] = status_log
        # log_list[7] = reason_log
        # print(log_list)
        # insert_log_into_table(cursor, log_list)
        # conn.commit()
        # log_list = [None] *10
        
        
    except Exception as e:
        print(f"An unexpected error occurred: {str(e)}")
        table_name_log = "rbi_odi"
        current_datetime = get_current_datetime()
        date_of_scraping = current_datetime
        total = get_data_count(cursor)
        log_list[0] = table_name_log
        log_list[10] = date_of_scraping
        status_log = "Failure"
        reason_log = " 404 error, website is not opened"
        log_list[1] = status_log
        log_list[4] = total 
        log_list[8] = reason_log
        print(log_list)
        insert_log_into_table(cursor, log_list)
        conn.commit()
        log_list = [None] *11
        sys.exit("error occured")

    finally:
        if 'browser' in locals() and browser is not None:
            browser.quit()




def extract_data_from_db(year,cursor):
    remaining_months = []
    # Your existing code here
    column_name = 'year'
    desired_value = year

    query = f"SELECT * FROM odi WHERE {column_name} = %s"
    cursor.execute(query, (desired_value,))

    # Fetch rows based on the condition
    selected_rows = cursor.fetchall()

    # Create a list to store the 10th values without duplicates
    months_in_database = set()
    year_database = set()

    # Extract the 10th values from each selected row and add to the list
    for row in selected_rows:
        if len(row) > 9:  
            months_in_database.add(row[10])
            year_database.add(row[11])


    # Print the unique 10th values
    print("These Months data are already in the database:", months_in_database)
    print("year",year_database)

    for a_month in actual_months:
        if a_month in months_in_database:
            print(a_month,"These month data is already inserted into the database")
        else:
            remaining_months.append(a_month)


    print(remaining_months,"These months are not in the database")
    if len(remaining_months) > 0:
        download_files_for_year(year,outer_month_xpath,download_xpath,cursor,remaining_months)
    else:
        print("no months want to be inserted into the database")
        print(remaining_months,"All the months are inserted for this year",year)

    pass



def insert_log_into_table(cursor, log_list):
    query = """
        INSERT INTO log_table (source_name, script_status, data_available, data_scraped, total_record_count, month, year, file_name, failure_reason, comments, date_of_scraping)
        VALUES (%(source_name)s, %(script_status)s, %(data_available)s, %(data_scraped)s, %(total_record_count)s, %(month)s, %(year)s, %(file_name)s, %(failure_reason)s, %(comments)s, %(date_of_scraping)s)
    """
    values = {
        'source_name': log_list[0] if log_list[0] else None,
        'script_status': log_list[1] if log_list[1] else None,
        'data_available': log_list[2] if log_list[2] else None,
        'data_scraped': log_list[3] if log_list[3] else None,
        'total_record_count': log_list[4] if log_list[4] else None,
        'month': log_list[5] if log_list[5] else None,
        'year': log_list[6] if log_list[6] else None,
        'file_name': log_list[7] if log_list[7] else None,
        'failure_reason': log_list[8] if log_list[8] else None,
        'comments': log_list[9] if log_list[9] else None,
        'date_of_scraping': log_list[10] if log_list[10] else None,
    }

    cursor.execute(query, values)







cursor.execute("SELECT * FROM odi")

# Fetch all rows
all_rows = cursor.fetchall()

years = []


def get_current_year():
    current_year = datetime.now().year
    years.append(current_year)
    return current_year


current_year = get_current_year()
previous_year = current_year -1
years.append(previous_year)
print("The current year is:",current_year)
print("The previous year is:",previous_year)



for year in years:
    extract_data_from_db(year,cursor)
else:
    table_name_log = "rbi_ecb"
    current_datetime = get_current_datetime()
    date_of_scraping = current_datetime
    log_list[0] = table_name_log
    log_list[10] = date_of_scraping
    incremental_count_total = get_data_count(cursor)
    if incremental_count_total <= total_count_first:
        status_log = "Success"
        comments = "No new data found"
        log_list[1] = status_log
        log_list[9] = comments
        total = get_data_count(cursor)
        log_list[4] = total       
        print(f"log for {date_of_scraping}")
        print(log_list)
        insert_log_into_table(cursor, log_list)
        conn.commit()
        log_list = [None] *11


conn.commit()
conn.close()











































