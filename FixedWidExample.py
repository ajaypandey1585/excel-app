import pandas as pd

# Specify the file path of your tab-delimited text file
file_path = 'newfile2.txt'

# Read the tab-delimited text file into a pandas DataFrame
# You need to specify the delimiter as '\t' for tab-delimited files
df = pd.read_csv(file_path, delimiter='\t')

# Export the DataFrame to an Excel file
excel_file_path = 'output_excel_file.xlsx'
df.to_excel(excel_file_path, index=False)

print("Data has been successfully imported into Excel.")
