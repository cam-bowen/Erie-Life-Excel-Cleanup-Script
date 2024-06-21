import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Paths to input and output files
input_csv = 'input_data.csv'
filtered_csv = 'filtered_output.csv'
output_xlsx = 'filtered_output.xlsx'

try:
    # Read CSV file
    df = pd.read_csv(input_csv)
    
    # Convert 'Unnamed: 8' column to datetime format (if necessary)
    df['Unnamed: 8'] = pd.to_datetime(df['Unnamed: 8'], errors='coerce')
    
    # Filter rows where 'Unnamed: 8' has a valid date
    filtered_df = df[df['Unnamed: 8'].notnull()].copy()  # Ensure to make a copy to avoid SettingWithCopyWarning
    
    # Drop specified columns
    columns_to_drop = ['Agent', 'Commissions Year (1)', 'Unnamed: 4', 
                       'Universal Life\nExcess', 'Universal Life',  # Corrected column name with newline
                       'Group Life', 'Annuity', 'Total', 'Total First Year and Renewal']
    
    # Ensure to strip any extra spaces from column names
    filtered_df.columns = filtered_df.columns.str.strip()
    
    # Drop columns by checking if they exist in the DataFrame
    filtered_df = filtered_df.drop(columns=[col for col in columns_to_drop if col in filtered_df.columns])
    
    # Rename columns for clarity
    filtered_df.rename(columns={'Unnamed: 1': 'Policy Number', 
                                'Unnamed: 2': 'Type',
                                'Traditional Life': 'Insured or Company Name',
                                'Unnamed: 6': 'Effective Dates',
                                'Unnamed: 8': 'Due Date',
                                'Unnamed: 10': 'Received Date',
                                'Unnamed: 11': 'Policy Year',
                                'Unnamed: 13': 'Basis',
                                'Unnamed: 14': 'Base Amount',
                                'Unnamed: 16': '%',
                                'Unnamed: 18': 'FYC',
                                'Unnamed: 20': 'Advanced Recovery',
                                'Unnamed: 21': 'Renewal'},
                       inplace=True)
    
    # Convert columns to numeric where necessary
    filtered_df['FYC'] = pd.to_numeric(filtered_df['FYC'], errors='coerce')
    filtered_df['Renewal'] = pd.to_numeric(filtered_df['Renewal'], errors='coerce')
    
    # Calculate commissions
    filtered_df['Commissions'] = filtered_df['FYC'] - filtered_df['Renewal']
    
    # Save filtered data to CSV
    filtered_df.to_csv(filtered_csv, index=False)
    print(f"Filtered data saved to {filtered_csv}")
    
    # Convert filtered DataFrame to Excel sheet
    wb = Workbook()
    ws = wb.active
    
    for r in dataframe_to_rows(filtered_df, index=False, header=True):
        ws.append(r)
    
    # Save Excel workbook to XLSX format
    wb.save(output_xlsx)
    print(f"Filtered data saved to {output_xlsx}")

except Exception as e:
    print(f"Error: {e}")

