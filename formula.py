import pandas as pd
import openpyxl

def add_formula_to_excel(input_excel, comparison_excel, output_excel):
    # Read the input Excel files
    df_input = pd.read_excel(input_excel)
    df_comparison = pd.read_excel(comparison_excel)
    
    # Merge the DataFrames on similar columns A to K
    merged_df = pd.merge(df_input, df_comparison, on=['BC', 'VALOR', 'TART', 'TYPE', 'BID-PRICE', 'BID-DATE', 'ASK-PRICE', 'ASK-DATE', 'VAL-PRICE', 'VAT', 'VAL-DATE'], how='left')
    
    # Apply the formula in the last column (column L)
    merged_df['L'] = merged_df.apply(lambda row: f'=IF((SUMIFS(E$3:E$52, H$3:H$52, {row["E"]})),"{row["E"]}", "Not Matching")', axis=1)
    
    # Write the updated DataFrame to the output Excel file
    merged_df.to_excel(output_excel, index=False)

# Example usage:
input_excel = 'output_file.xlsx'  # Replace 'output_file.xlsx' with your first Excel file name
comparison_excel = 'output_file2.xlsx'  # Replace 'output_file2.xlsx' with your second Excel file name
output_excel = 'output_combined.xlsx'  # Replace 'output_combined.xlsx' with your output Excel file name

add_formula_to_excel(input_excel, comparison_excel, output_excel)
