import pandas as pd
import os

# Ask user for folder path
folder_path = input("Paste the folder path containing CSV files and press Enter: ").strip()

# Output file path (same folder)
output_file = os.path.join(folder_path, "All_CSV_Converted_Into_Sheets.xlsx")

# Create Excel writer
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    
    # Loop through all files in the folder
    for file in os.listdir(folder_path):
        if file.lower().endswith(".csv"):
            file_path = os.path.join(folder_path, file)
            
            # Read CSV
            df = pd.read_csv(file_path)
            
            # Sheet name = file name without extension
            sheet_name = os.path.splitext(file)[0][:31]  # Excel sheet name limit
            
            # Write to Excel
            df.to_excel(writer, sheet_name=sheet_name, index=False)

print("\nAll CSV files successfully converted into Excel sheets!")
print(f"File saved at: {output_file}")


