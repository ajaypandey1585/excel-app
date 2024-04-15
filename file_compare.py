import pandas as pd

def compare_and_mark_unmatched(input_excel, output_excel, comparison_excel, column_name):
    # Read the input Excel files, ignoring the top row
    df_input = pd.read_excel(input_excel, header=1)
    df_comparison = pd.read_excel(comparison_excel, header=1)
    
    # Check if the specified column exists in both DataFrames
    if column_name not in df_input.columns or column_name not in df_comparison.columns:
        print(f"Column '{column_name}' not found in one of the input Excel files.")
        return
    
    # Merge the DataFrames on the specified column
    merged_df = pd.merge(df_input, df_comparison, on=column_name, how='left', suffixes=('_input', '_comparison'))
    
    # Mark the unmatched rows
    merged_df['Unmatched'] = 'Match'
    unmatched_indices = merged_df[merged_df[column_name].isnull()].index
    merged_df.loc[unmatched_indices, 'Unmatched'] = 'Unmatched'
    
    # Write the updated DataFrame to the output Excel file
    merged_df.to_excel(output_excel, index=False)

# Example usage:
input_excel = 'output_file.xlsx'  # Input Excel file name
output_excel = 'ComparedExcel.xlsx'  # Output Excel file name
comparison_excel = 'output_file2.xlsx'  # Comparison Excel file name
column_name = 'BC   '  # Dynamic column name

compare_and_mark_unmatched(input_excel, output_excel, comparison_excel, column_name)
