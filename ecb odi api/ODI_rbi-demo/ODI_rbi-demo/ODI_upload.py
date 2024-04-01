import os
import glob
import pandas as pd
from datetime import datetime
import mysql.connector
from sqlalchemy import create_engine
import re

output_folder_combined = r'C:\Users\surya.8263\Desktop\ODI\combined_files'
os.makedirs(output_folder_combined, exist_ok=True)

output_folder_outside = r'C:\Users\surya.8263\Desktop\ODI\combinedata'
os.makedirs(output_folder_outside, exist_ok=True)

main_folder_path = r'C:\Users\surya.8263\Desktop\ODI\years'
year_folders = [f for f in os.listdir(main_folder_path) if os.path.isdir(os.path.join(main_folder_path, f))]



# type1_columns = {'Unnamed: 0': 'SI_NO', 'Unnamed: 1': 'Name_of_the_Indian_Party','Unnamed: 2': 'Name_of_the_JV_WOS', 'Unnamed: 3': 'Whether_JV_WOS','Unnamed: 4': 'Overseas_Country', 'Unnamed: 6': 'Major_Activity','Unnamed: 7': 'FC_Equity', 'Unnamed: 8': 'FC_Loan', 'Unnamed: 9':'FC_Guarantee_Issued','Unnamed: 10':'FC_Total','Month':'Month','Year':'Year'}
# type2_columns = {'Foreign Exchange Department': 'SI_NO', 'Unnamed: 1': 'Name_of_the_Indian_Party','Unnamed: 2': 'Name_of_the_JV_WOS', 'Unnamed: 3': 'Whether_JV_WOS','Unnamed: 4': 'Overseas_Country', 'Unnamed: 6': 'Major_Activity','Unnamed: 7': 'FC_Equity', 'Unnamed: 8': 'FC_Loan', 'Unnamed: 9':'FC_Guarantee_Issued','Unnamed: 10':'FC_Total','Month':'Month','Year':'Year'}
# type3_columns = {'Sr.': 'SI_NO', 'Name of the Indian Party': 'Name_of_the_Indian_Party','Name of the JV/WOS': 'Name_of_the_JV_WOS', 'Whether JV/WOS': 'Whether_JV_WOS','Overseas Country': 'Overseas_Country', 'Major Activity': 'Major_Activity','Financial Commitment': 'FC_Equity', 'Unnamed: 7': 'FC_Loan', 'Unnamed: 8':'FC_Guarantee_Issued','Unnamed: 9':'FC_Total','Month':'Month','Year':'Year'}
# type4_columns = {'Unnamed: 1': 'SI_NO', 'Unnamed: 2': 'Name_of_the_Indian_Party','Unnamed: 3': 'Name_of_the_JV_WOS', 'Unnamed: 4': 'Whether_JV_WOS','Unnamed: 5': 'Overseas_Country', 'Unnamed: 7': 'Major_Activity','Unnamed: 8': 'FC_Equity', 'Unnamed: 9': 'FC_Loan', 'Unnamed: 10':'FC_Guarantee_Issued','Unnamed: 11':'FC_Total','Month':'Month','Year':'Year'}

common_columns = ['SI_NO', 'Name_of_the_Indian_Party', 'Name_of_the_JV_WOS', 'Whether_JV_WOS', 'Overseas_Country',
                  'Major_Activity', 'FC_Equity', 'FC_Loan', 'FC_Guarantee_Issued', 'FC_Total', 'Month', 'Year']

combined_data = []  # Reset combined_data as it's now a list

for year_folder in year_folders:
    year_path = os.path.join(main_folder_path, year_folder)
    month_folders = [f for f in os.listdir(year_path) if os.path.isdir(os.path.join(year_path, f))]

    for month_folder in month_folders:
        month_path = os.path.join(year_path, month_folder)

        excel_files = glob.glob(os.path.join(month_path, '*.xlsx')) + glob.glob(os.path.join(month_path, '*.xls*'))

        for file in excel_files:
            if not file.startswith('~$'):
                try:
                    xls = pd.ExcelFile(file)

                    # Read only the first sheet
                    df = xls.parse(sheet_name=0)

                    # Remove empty columns
                    df = df.dropna(axis=1, how='all')

                    # Remove rows with four or more empty cells
                    df = df.dropna(thresh=df.shape[1]-2)

                    month_year = month_folder.split(' ')
                    if len(month_year) == 1:
                        month = month_year[0]
                    else:
                        filename_parts = os.path.splitext(os.path.basename(file))[0].split('for')
                        month_year = filename_parts[-1].split(' ')
                        if len(month_year) > 1:
                            month = month_year[1]

                    # Extract year from the file name using regular expressions
                    year_match = re.search(r'(\d{4})', month_folder)
                    year = year_match.group(1) if year_match else ''

                    # Add new columns to the DataFrame
                    df['Month'] = month
                    df['Year'] = year

                    # Remove any remaining empty columns
                    df = df.dropna(axis=1, how='all')

                    df.columns = common_columns

                    # Reorder columns to have 'Month' and 'Year' at the end
                    column_order = [col for col in df.columns if col not in ['Month', 'Year']] + ['Month', 'Year']

                    df = df[column_order]
                    # print(df.columns)

                    output_file_path = os.path.join(output_folder_combined, f'{year_folder}_{month_folder}_modified.xlsx')
                    df.to_excel(output_file_path, index=False)
                    print(f"Modified data saved to {output_file_path}")

                    combined_data.append(output_file_path)
                    print(f"Processed and saved combined columns: {output_file_path}")

                except Exception as e:
                    print(f"Error reading {file}: {e}")

# After reading and combining the Excel files into the combined_data list
if combined_data:
    all_dataframes = [pd.read_excel(file) for file in combined_data]
    combined_df = pd.concat(all_dataframes, ignore_index=True)
    specified_columns = ['SI_NO', 'Name_of_the_Indian_Party', 'Name_of_the_JV_WOS', 'Whether_JV_WOS', 'Overseas_Country', 'Major_Activity', 'FC_Equity', 'FC_Loan', 'FC_Guarantee_Issued', 'FC_Total', 'Month', 'Year', 'date_scraped']
    columns_to_check_duplicates = ['Name_of_the_Indian_Party', 'Name_of_the_JV_WOS', 'Whether_JV_WOS', 'Overseas_Country', 'Major_Activity', 'FC_Equity', 'FC_Loan', 'FC_Guarantee_Issued', 'FC_Total', 'Month', 'Year']

    # Apply the filtering conditions to the DataFrame
    exclude_conditions = (
        (combined_df['Month'] == 'December') & (combined_df['Year'] == '2011') |
        (combined_df['Month'] == 'November') & (combined_df['Year'] == '2012')
    )

    # Apply the filter
    combined_df = combined_df[~exclude_conditions]

    existing_columns = set(combined_df.columns)
    specified_columns = [
        col for col in specified_columns if col in existing_columns
    ]

    if specified_columns:
        combined_df = combined_df[specified_columns]

        # Remove rows where any of the specified columns are empty
        combined_df.dropna(subset=['Name_of_the_Indian_Party', 'Name_of_the_JV_WOS', 'Whether_JV_WOS', 'Overseas_Country', 'Major_Activity', 'FC_Equity', 'FC_Loan', 'FC_Guarantee_Issued', 'FC_Total', 'Month', 'Year'], how='any', inplace=True)

        combined_df.drop_duplicates(subset=columns_to_check_duplicates, inplace=True)

        output_file_path_combined = os.path.join(output_folder_outside, 'combined_data1.xlsx')
        combined_df.to_excel(output_file_path_combined, index=False, engine='openpyxl')
    else:
        print("No matching columns found in the combined DataFrame.")
else:
    print("No data to combine.")


# MySQL Database Configuration
username = 'root'
password = 'root'
database_name = 'oversease_direct_investment'
hostname = '127.0.0.1'

column_data_types = {
    'SI_NO': 'int NOT NULL AUTO_INCREMENT',
    'Name_of_the_Indian_Party': 'VARCHAR(200)',
    'Name_of_the_JV_WOS': 'VARCHAR(200)',
    'Whether_JV_WOS': 'VARCHAR(20)',
    'Overseas_Country': 'VARCHAR(250)',
    'Major_Activity': 'VARCHAR(250)',
    'FC_Equity': 'FLOAT',  
    'FC_Loan': 'FLOAT',
    'FC_Guarantee_Issued': 'FLOAT',
    'FC_Total': 'FLOAT',
    'Month': 'VARCHAR(20)',
    'Year': 'VARCHAR(10)',
    'date_scraped': 'TIMESTAMP',
}

table_name = 'rbi_odi_output'

def create_table():
    conn = mysql.connector.connect(
        host='127.0.0.1',
        user='root',
        password='root',
        database='oversease_direct_investment'
    )
    cursor = conn.cursor()

    columns_sql = ', '.join([f"`{col}` {dtype}" for col, dtype in column_data_types.items() if col != 'SI_NO'])

    create_table_query = f"CREATE TABLE IF NOT EXISTS `{table_name}` (Sr_No int NOT NULL AUTO_INCREMENT PRIMARY KEY, {columns_sql})"
    cursor.execute(create_table_query)
    conn.commit()
    conn.close()

def insert_into_table():
    combined_data = pd.read_excel(output_file_path_combined)
    
    combined_data = combined_data[~combined_data['SI_NO'].astype(str).str.startswith('Sr.')]


    # Rename columns as needed
    combined_data = combined_data.rename(columns={
        'SI_NO': 'SI_NO',
        'Name_of_the_Indian_Party': 'Name_of_the_Indian_Party',
        'Name_of_the_JV_WOS': 'Name_of_the_JV_WOS',
        'Whether_JV_WOS': 'Whether_JV_WOS',
        'Overseas_Country': 'Overseas_Country',
        'Major_Activity': 'Major_Activity',
        'FC_Equity': 'FC_Equity',
        'FC_Loan': 'FC_Loan',
        'FC_Guarantee_Issued': 'FC_Guarantee_Issued',
        'FC_Total': 'FC_Total',
        'Month': 'Month',
        'Year': 'Year',
        'date_scraped': 'date_scraped'
    })

    count_by_year_month = combined_data.groupby(['Year', 'Month']).size().reset_index(name='Count')
    print("Count by Year and Month:")
    print(count_by_year_month)

    combined_data['Year'] = combined_data['Year'].astype(str)
    combined_data['Year'] = combined_data['Year'].str.replace(',', '')

    current_time = datetime.now().replace(microsecond=0)
    combined_data['date_scraped'] = current_time

    combined_data = combined_data.drop('SI_NO', axis=1)

    engine = create_engine(f'mysql+mysqlconnector://{username}:{password}@{hostname}/{database_name}')
    combined_data = combined_data.where(pd.notnull(combined_data), None)
    
    combined_data = combined_data[~exclude_conditions]

    combined_data.to_sql(table_name, engine, if_exists='append', index=False)

create_table()
insert_into_table()
print("Successfully Uploaded to Database.")