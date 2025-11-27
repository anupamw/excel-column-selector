# Excel Column Selector

A Python CLI tool to interactively select columns from an Excel file and save a filtered version.

## Features

- Read Excel files (.xlsx, .xls)
- Interactive column selection using checkboxes
- Save filtered Excel file with selected columns only
- Preserves original data types and formatting

## Installation

```bash
# Create a virtual environment (recommended)
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Usage

```bash
python column_selector.py <path_to_excel_file>
```

### Example

```bash
python column_selector.py /path/to/data.xlsx
```

The program will:
1. Read the Excel file
2. Display an interactive list of all available columns
3. Allow you to select columns using:
   - Arrow keys to navigate
   - Spacebar to select/deselect
   - Enter to confirm selection
4. Save the filtered file as `data_filtered.xlsx` in the same directory

## Output

The filtered Excel file will be saved in the same directory as the original file with `_filtered` appended to the filename.

Example: `data.xlsx` → `data_filtered.xlsx`

## Requirements

- Python 3.7+
- pandas
- openpyxl
- inquirer

