# Excel SRID Data Processing and Contractor Update Tool

## Overview
This is a Python-based tool designed to automate the processing and updating of a master Excel file based on **Service Request ID (SRID)** values. It merges specific sheets, clears and updates the **"CONTRACTOR"** and **"FASTX"** columns, and generates a new output file. The program ensures efficient handling of contractor and status data across multiple sheets.

## Features
- Prompts the user to select two input Excel files and an output directory.
- Merges the following sheets into one:
  - "Ανατεθειμένα για κατασκευή"
  - "Ανατεθειμένες αυτοψίες"
  - "Εντολές στο ίδιο BID"
- Clears and updates the **"CONTRACTOR"** (or **"contractor"**) columns based on SRID values from the second input file.
- Copies **"FASTX"** values into **"CONTRACTOR"** columns before updating them.
- Updates the **"FASTX"** column based on the SRID-matching **"Κατάσταση"** values from the second input file.
- Processes all other sheets while skipping unnecessary ones like "Βλάβες" and "Pivots."
- Saves the output Excel file in the selected output directory.

## Requirements
- Python 3.x
- pandas
- xlsxwriter
- tkinter

## Installation
Clone this repository or download the script files:

```bash
git clone https://github.com/pbonotis15/excel_srid_contractor_update_tool.git
cd your-repo-name
```
Install the required Python libraries:

```bash
pip install pandas xlsxwriter
```
## Usage
Run the SRID_contractor_update.py script:

```bash
python SRID_contractor_update.py
```

When prompted:
- Select the first input Excel file (master file containing SRIDs).
- Select the second input Excel file (containing the SRID updates).
- Select the output directory where the final processed Excel file will be saved.

The script will:
- Merge specified sheets into one.
- Clear and update the "CONTRACTOR" (and "contractor") and "FASTX" columns across all sheets.
- Save the updated file to the specified output folder.

## Contributing
If you'd like to contribute to this project, please fork the repository and use a feature branch. Pull requests are warmly welcomed.
