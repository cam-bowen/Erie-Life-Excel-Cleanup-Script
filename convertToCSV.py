import pandas as pd

# Replace 'input_file.xlsx' with your actual Excel file path
input_file = 'ERIE_Life_EE3012_06.14.24.xlsx'
output_csv = 'input_data.csv'

try:
    # Read Excel file
    df = pd.read_excel(input_file)
    
    # Convert and save to CSV
    df.to_csv(output_csv, index=False)
    
    print(f"Excel file '{input_file}' converted to CSV: '{output_csv}'")

except Exception as e:
    print(f"Error: {e}")
