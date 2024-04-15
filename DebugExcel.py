def text_to_excel_fixed_width(input_file, output_file, encoding):
    try:
        print("Converting fixed-width file...")
        print("Input file:", input_file)
        print("Output file:", output_file)
        print("Encoding:", encoding)
        # Read the entire file to determine the structure
        with open(input_file, 'r', encoding=encoding) as file:
            lines = [line.rstrip('\n') for line in file]
        
        # Assuming the structure is determined here (e.g., column positions)
        # Perform fixed-width conversion accordingly
        
        # Write to Excel
        # df.to_excel(output_file, index=False)
        
        return True
    except Exception as e:
        print("Error during fixed-width conversion:", str(e))
        return False


def text_to_excel_delimited(input_file, output_file, delimiter, encoding):
    try:
        print("Converting delimited file...")
        print("Input file:", input_file)
        print("Output file:", output_file)
        print("Encoding:", encoding)
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
        df = pd.DataFrame(lines[1:], columns=header)
        
        # Write to Excel
        df.to_excel(output_file, index=False)
        
        return True
    except Exception as e:
        print("Error during delimited conversion:", str(e))
        return False


def text_to_excel_unknown(input_file, output_file, encoding):
    try:
        print("Converting unknown file type...")
        print("Input file:", input_file)
        print("Output file:", output_file)
        print("Encoding:", encoding)
        # Read the entire file to determine the structure
        with open(input_file, 'r', encoding=encoding) as file:
            lines = [line.rstrip('\n') for line in file]
        
        # Assuming the structure is determined here (e.g., by heuristic)
        # Perform conversion accordingly
        
        # Write to Excel
        # df.to_excel(output_file, index=False)
        
        return True
    except Exception as e:
        print("Error during conversion:", str(e))
        return False
