import pandas as pd
from tkinter import Tk, filedialog

# Function to ask for file selection
def ask_for_file(prompt_text):
    print(prompt_text)
    Tk().withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename()  # Ask user for the file path
    print(f"Selected file: {file_path}")
    return file_path

# Function to ask for folder selection
def ask_for_folder(prompt_text):
    print(prompt_text)
    Tk().withdraw()  # Hide the root window
    folder_path = filedialog.askdirectory()  # Ask user for the folder path
    print(f"Selected folder: {folder_path}")
    return folder_path

# Step 1: Ask user to select the first input file
file_path = ask_for_file("Please select the first input Excel file (e.g., 'unique_tasks_12-10-2024.xlsx'):")

# Load the input Excel file
xls = pd.ExcelFile(file_path)

# Step 2: Ask user to select the second input file
input2_file_path = ask_for_file("Please select the second input Excel file (e.g., 'ifs-fsm-formatted_12-10-2024.xlsx'):")

# Load the second input file (input2 file)
input2_df = pd.read_excel(input2_file_path)

# Step 3: Ask user to select the output folder
output_folder = ask_for_folder("Please select the folder where the output file will be saved:")

# Define the relevant columns from input2 file
srid_mapping = input2_df[['SR ID', 'Όνομα', 'Κατάσταση']]

# Load the specific sheets to merge
sheet_1 = pd.read_excel(xls, sheet_name="Ανατεθειμένα για κατασκευή")
sheet_2 = pd.read_excel(xls, sheet_name="Ανατεθειμένες αυτοψίες")
sheet_3 = pd.read_excel(xls, sheet_name="Εντολές στο ίδιο BID")

# Merge the specified sheets into one
merged_sheet = pd.concat([sheet_1, sheet_2, sheet_3])

# Create a new Excel writer for the output file
output_file_path = f"{output_folder}/output_excel_with_modifications.xlsx"
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    # Save the merged sheet with the new name
    merged_sheet.to_excel(writer, sheet_name="ΑΥΤΟΨΙΕΣ || ΚΑΤΑΣΚΕΥΕΣ || BID", index=False)

    # Pass all other sheets except for "Βλάβες" and "Pivots"
    for sheet in xls.sheet_names:
        if sheet == "Βλάβες" or sheet == "Pivots":  # Skip "Βλάβες" and "Pivots" sheets
            continue
        elif sheet not in ["Ανατεθειμένα για κατασκευή", "Ανατεθειμένες αυτοψίες", "Εντολές στο ίδιο BID"]:
            pd.read_excel(xls, sheet_name=sheet).to_excel(writer, sheet_name=sheet, index=False)

# Reload the output file for modifications
output_xls = pd.ExcelFile(output_file_path)

# Step 4: Clear "CONTRACTOR", "contractor", copy "FASTX" into it, clear "FASTX", and perform SRID update
updated_sheets = {}

for sheet in output_xls.sheet_names:
    df = pd.read_excel(output_xls, sheet_name=sheet)
    
    # Check if the sheet contains "CONTRACTOR" or "contractor" column
    if 'CONTRACTOR' in df.columns or 'contractor' in df.columns:
        # Step 1: Clear both "CONTRACTOR" and "contractor"
        if 'CONTRACTOR' in df.columns:
            df['CONTRACTOR'] = ''
        if 'contractor' in df.columns:
            df['contractor'] = ''
        
        # Step 2: If "FASTX" exists, copy "FASTX" values to "CONTRACTOR" and "contractor"
        if 'FASTX' in df.columns:
            if 'CONTRACTOR' in df.columns:
                df['CONTRACTOR'] = df['FASTX']
            if 'contractor' in df.columns:
                df['contractor'] = df['FASTX']
        
        # Step 3: Clear the "FASTX" column if it exists
        if 'FASTX' in df.columns:
            df['FASTX'] = ''
        
        # Step 4: Merge with SRID data to update CONTRACTOR, contractor, and FASTX
        df_updated = df.merge(srid_mapping, how='left', left_on='SR ID', right_on='SR ID')
        
        # Update CONTRACTOR and contractor where SRID matches with "Όνομα"
        if 'CONTRACTOR' in df_updated.columns:
            df_updated['CONTRACTOR'] = df_updated['Όνομα'].combine_first(df_updated['CONTRACTOR'])
        if 'contractor' in df_updated.columns:
            df_updated['contractor'] = df_updated['Όνομα'].combine_first(df_updated['contractor'])
        
        # Update FASTX where SRID matches with "Κατάσταση"
        if 'FASTX' in df_updated.columns:
            df_updated['FASTX'] = df_updated['Κατάσταση'].combine_first(df_updated['FASTX'])
        
        # Drop helper columns from merge
        df_updated = df_updated.drop(columns=['Όνομα', 'Κατάσταση'])
    else:
        df_updated = df  # No update needed, keep as is
    
    # Add the updated sheet to the dictionary
    updated_sheets[sheet] = df_updated

# Step 5: Save the final output file
final_output_file_path = f"{output_folder}/final_output_without_pivot.xlsx"
with pd.ExcelWriter(final_output_file_path, engine='xlsxwriter') as writer:
    for sheet_name, updated_df in updated_sheets.items():
        updated_df.to_excel(writer, sheet_name=sheet_name, index=False)

# Output file is saved
print(f"Final output saved at: {final_output_file_path}")