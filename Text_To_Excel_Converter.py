import datetime
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
import pandas as pd
import os

def text_to_excel_fixed_width(input_file, output_file, encoding):
    try:
        # Read the entire file to determine the structure
        with open(input_file, 'r', encoding=encoding) as file:
            lines = [line.rstrip('\n') for line in file]
        
        # Assuming the structure is determined here (e.g., column positions)
        # Perform fixed-width conversion accordingly
        
        #Write to Excel
        df = pd.DataFrame(lines[1:], columns=lines[0])
        df.to_excel(output_file, index=False)
        
        return True
    except Exception as e:
        print(f"Error during fixed-width conversion: {str(e)}")
        return False

def text_to_excel_delimited(input_file, output_file, delimiter, encoding):
    try:
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
        print(f"Error during delimited conversion: {str(e)}")
        return False

def detect_file_type(input_file):
    try:
        with open(input_file, 'r', encoding='utf-8') as file:
            # Read the first few lines to determine the structure
            lines = [line.rstrip('\n') for line in file]
        
        # Check if all lines have the same length
        line_lengths = set(len(line) for line in lines)
        if len(line_lengths) == 1:
            return 'fixed_width'  # All lines have the same length -> Fixed-width
        
        # Check for common delimiters in the first few lines
        delimiters = [',', ';', '|', '\t']
        for delimiter in delimiters:
            if all(delimiter in line for line in lines):
                return 'delimited'  # Common delimiter found -> Delimited
        
        # Default to unknown if none of the above conditions are met
        return 'unknown'
    except Exception as e:
        print(f"Error detecting file type: {str(e)}")
        return None

def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    entry_file.delete(0, tk.END)
    entry_file.insert(0, filename)

def convert_to_excel():
    input_file = entry_file.get()
    if not input_file:
        lbl_status.config(text="Please select a file.")
        return
    
    # Get the current date and time
    current_datetime = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    
    # Append the current date and time to the output file name
    output_file_name = f"{os.path.splitext(os.path.basename(input_file))[0]}_{current_datetime}.xlsx"
    output_file = os.path.join(os.path.dirname(input_file), output_file_name)        
    
    # Detect file type
    file_type = detect_file_type(input_file)
    
    if file_type == 'fixed_width':
        # Perform fixed-width conversion
        success = text_to_excel_fixed_width(input_file, output_file, 'utf-8')
    elif file_type == 'delimited':
        # Default to unknown conversion
        success = text_to_excel_delimited(input_file, output_file, ',', 'utf-8')
    else:
        # Default to unknown conversion
        success = False
    
    if success:
        lbl_status.config(text=f"Conversion completed successfully. Output file: {output_file}")
    else:
        lbl_status.config(text="Error during conversion.")

# Create Tkinter window
window = tk.Tk()
window.title("Text to Excel Converter")
# Add logo to the title bar
window.iconbitmap('favicon.ico')
window.geometry("400x250")

# Load logo image
logo_image = Image.open("six-logo.png")
logo_image = logo_image.resize((175, 50))  # Resize the logo
logo_photo = ImageTk.PhotoImage(logo_image)


# Create UI components
lbl_logo = tk.Label(window, image=logo_photo)
lbl_logo.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

lbl_file = tk.Label(window, text="Select a .txt file:")
lbl_file.grid(row=1, column=0, padx=10, pady=5)

entry_file = tk.Entry(window, width=30)
entry_file.grid(row=1, column=1, padx=10, pady=5)

btn_browse = tk.Button(window, text="Browse", command=browse_file)
btn_browse.grid(row=1, column=2, padx=10, pady=5)

btn_convert = tk.Button(window, text="Convert to Excel", command=convert_to_excel)
btn_convert.grid(row=2, column=1, padx=10, pady=5)

lbl_status = tk.Label(window, text="")
lbl_status.grid(row=3, column=0, columnspan=3, padx=10, pady=5)

# Run the Tkinter event loop
window.mainloop()
