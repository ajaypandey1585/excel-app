import pandas as pd

def text_to_excel(input_file, output_file, delimiter, encoding):
    # Read the first row to determine the number of columns and headers
    with open(input_file, 'r', encoding=encoding) as file:
        first_row = file.readline().rstrip('\n').split(delimiter)
        header = first_row if any(c.isalpha() for c in first_row) else None
    
    # Read the text file and filter out lines with extra delimiters
    lines = []
    with open(input_file, 'r', encoding=encoding) as file:
        for line in file:
            split_line = line.rstrip('\n').split(delimiter)
            if len(split_line) == len(first_row):
                lines.append(split_line)

    # Create DataFrame
    df = pd.DataFrame(lines, columns=header)

    # Write to Excel
    df.to_excel(output_file, index=False)

# Example usage:
input_file = 'newfile1.txt'  # Replace 'newfile1.txt' with your actual input file name
output_file = 'output_file.xlsx'
delimiter = ''  # You can change this to any delimiter you want
encoding = 'latin-1'  # Specify the desired encoding here

text_to_excel(input_file, output_file, delimiter, encoding)
