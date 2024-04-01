import os
import glob
import pandas as pd
import re
from datetime import datetime
import mysql.connector
from sqlalchemy import create_engine
from collections import defaultdict


output_folder_combined = r'C:\Users\surya.8263\Desktop\ODI\combined_files'
os.makedirs(output_folder_combined, exist_ok=True)

output_folder_outside = r'C:\Users\surya.8263\Desktop\ODI\combinedata'
os.makedirs(output_folder_outside, exist_ok=True)


main_folder_path = r'C:\Users\surya.8263\Desktop\ODI\years'
year_folders = [f for f in os.listdir(main_folder_path) if os.path.isdir(os.path.join(main_folder_path, f))]

combined_data = []


rows_to_delete = [
    'Grand Total',
    '* As reported by Authorized Dealers.',
    'Run Date:',
    'Page 1 of 1',
    '* As reported by Authorized Dealers in Form ODI . The amount reported towards Equity and loan represents the actual outflows during the month.',
    'Note:Classification and grouping of activity is as per NIC Code 1987 siadipp.nic.in/policy/nic/nic.ht',
    '* As reported by Authorized Dealers in Form ODI.'
    # Add any other text patterns you want to remove here
]
file_counts = defaultdict(int)
                    
for year_folder in year_folders:
    year_path = os.path.join(main_folder_path, year_folder)
    month_folders = [f for f in os.listdir(year_path) if os.path.isdir(os.path.join(year_path, f))]

    for month_folder in month_folders:                      
        month_path = os.path.join(year_path, month_folder)

        excel_files = glob.glob(os.path.join(month_path, '*.xlsx')) + glob.glob(os.path.join(month_path, '*.xls*'))

        for file in excel_files:
            if not file.startswith('~$'):  
                try:
                    df = pd.read_excel(file)

                   
                    for pattern in rows_to_delete:
                        df = df[~df.astype(str).apply(lambda row: row.str.contains(re.escape(pattern))).any(axis=1)]

                    
                    sr_rows = df[df.apply(lambda row: row.astype(str).str.startswith('Sr.')).any(axis=1)]
                    
                    
                    if not sr_rows.empty:
                        sr_row_index = sr_rows.index[0]
                        df.columns = df.iloc[sr_row_index]
                        df = df.iloc[sr_row_index + 1:]
                        
                        
                        df = df.dropna(axis=1, how='all')

                       
                        df = df.dropna(axis=0, how='all')

                        
                        selected_columns = df.iloc[:, :6]

                        
                        remaining_columns = df.iloc[:, 6:]
                        
                        
                        remaining_columns = remaining_columns[~remaining_columns.iloc[:, 0].fillna('').astype(str).str.contains('Financial')]
                        
                        
                        remaining_columns.columns = remaining_columns.iloc[0]
                        remaining_columns = remaining_columns.iloc[1:]

                        # Rename columns
                        new_column_names = {
                            'Equity *': 'FC_Equity *',
                            'Loan *': 'FC_Loan *',
                            'Guarantee Issued *': 'FC_Guarantee Issued *',
                            'Total': 'FC_Total'
                            
                        }
                        remaining_columns = remaining_columns.rename(columns=new_column_names)

                        
                        combined_df = pd.concat([selected_columns, remaining_columns], axis=1)
                        combined_df = combined_df.iloc[1:, :]


                    month_year = month_folder.split(' ')
                    if len(month_year) == 1:  
                        month = month_year[0]
                    else:  
                        filename_parts = os.path.splitext(os.path.basename(file))[0].split('for')
                        month_year = filename_parts[-1].split(' ')  
                        if len(month_year) > 1:  
                            month = month_year[1]

                    

                    year = re.sub(r'[^0-9]', '', year_folder)

                    
                    if 'Month' in combined_df.columns:
                        
                        combined_df['Month'] = month
                    else:
                       
                        combined_df['Month'] = month

                    
                    if 'Year' in combined_df.columns:
                        
                        combined_df['Year'] = year
                    else:
                        
                        combined_df['Year'] = year

                    
                    column_order = [col for col in combined_df.columns if col not in ['Month', 'Year']] + ['Month', 'Year']
                    combined_df = combined_df[column_order]

                    if 'Year' in combined_df.columns and 'Month' in combined_df.columns:
                        file_counts[(year, month)] += 1
                        
                    output_file_path_combined = os.path.join(output_folder_combined, f'combined_{os.path.basename(file)}')
                    combined_df.to_excel(output_file_path_combined, index=False, engine='openpyxl')
                                                                            
                    
                    combined_data.append(output_file_path_combined)
                    # print(combined_df.columns,'columns names')

                    print(f"Processed and saved combined columns: {output_file_path_combined}")
                except PermissionError:
                    print(f"Permission denied for file: {file}. Skipping this file.")
        # print("values :")
        # print(combined_df)
print("File counts by Year and Month:")
for key, value in file_counts.items():
    print(f"Year: {key[0]}, Month: {key[1]}, File Count: {value}")
    
if combined_data:
    all_dataframes = [pd.read_excel(file) for file in combined_data]
    combined_df = pd.concat(all_dataframes, ignore_index=True)
    specified_columns = ['Sr.','Name of the Indian Party', 'Name of the JV/WOS', 'Whether JV/WOS', 'Overseas Country', 'Major Activity', 'FC_Equity *', 'FC_Loan *', 'FC_Guarantee Issued *', 'FC_Total', 'Month', 'Year','Date_of_Scraping']  
    columns_to_check_duplicates = ['Name of the Indian Party','Name of the JV/WOS','Whether JV/WOS','Overseas Country','Major Activity','FC_Equity *','FC_Loan *','FC_Guarantee Issued *','FC_Total', 'Month', 'Year']
    if not combined_df.empty:
       
        existing_columns = set(combined_df.columns)
        specified_columns = [
            col for col in specified_columns if col in existing_columns
        ]

        if specified_columns:
           
            combined_df = combined_df[specified_columns]

            # Remove rows where any of the specified columns are empty
            combined_df.dropna(subset=['Name of the Indian Party','Name of the JV/WOS','Whether JV/WOS','Overseas Country','Major Activity','FC_Equity *','FC_Loan *','FC_Guarantee Issued *','FC_Total', 'Month', 'Year'], how='any', inplace=True)

            combined_df.drop_duplicates(subset=columns_to_check_duplicates, inplace=True)

            output_file_path_combined = os.path.join(output_folder_outside, 'combined_data1.xlsx')
            combined_df.to_excel(output_file_path_combined, index=False, engine='openpyxl')
        else:
            print("No matching columns found in the combined DataFrame.")
    else:
        print("No data to combine.") 


username = 'root1'
password = 'Mysql1234$'
database_name = 'rbi_odi'
hostname = '4.213.77.165'

# username = 'root'
# password = 'root'
# database_name = 'oversease_direct_investment'
# hostname = '127.0.0.1'

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
        host='4.213.77.165',
        user='root1',
        password='Mysql1234$',
        database='rbi_odi'
    )
    cursor = conn.cursor()
    
    columns_sql = ', '.join([f"`{col}` {dtype}" for col, dtype in column_data_types.items() if col != 'SI_NO'])
    
    create_table_query = f"CREATE TABLE IF NOT EXISTS `{table_name}` (Sr_No int NOT NULL AUTO_INCREMENT PRIMARY KEY, {columns_sql})"
    cursor.execute(create_table_query)
    conn.commit()
    conn.close()



    
def insert_into_table():
    combined_data = pd.read_excel(output_file_path_combined)
    
    # Rename columns as needed
    combined_data = combined_data.rename(columns={
        'Sr.' : 'SI_NO',
        'Name of the Indian Party': 'Name_of_the_Indian_Party',
        'Name of the JV/WOS': 'Name_of_the_JV_WOS',
        'Whether JV/WOS': 'Whether_JV_WOS',
        'Overseas Country':'Overseas_Country',
        'Major Activity':'Major_Activity',
        'FC_Equity *':'FC_Equity',
        'FC_Loan *':'FC_Loan',
        'FC_Guarantee Issued *':'FC_Guarantee_Issued',
        'FC_Total':'FC_Total',
        'Month':'Month',
        'Year':'Year',
        'Date_of_Scraping':'date_scraped',
        
        
        
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
    combined_data.to_sql(table_name, engine, if_exists='append', index=False)
    print("Successfully Upload In DataBase...")
    
    # conn = mysql.connector.connect(
    #     host='4.213.77.165',
    #     user='root1',
    #     password='Mysql1234$',
    #     database='rbi_odi'
    # )
    # cursor = conn.cursor()

    # count_query = "SELECT Year, COUNT(*) AS row_count FROM rbi_odi_output GROUP BY Year"
    # cursor.execute(count_query)
    # rows = cursor.fetchall()

    # for row in rows:
    #     year = row[0]
    #     row_count = row[1]
    #     print(f"Year: {year}, Row Count: {row_count}")

    # conn.close()

create_table()
insert_into_table()