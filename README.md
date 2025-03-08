# Excel Analyzer

This repository contains a PowerShell script that automates the analysis of Excel files. It uses Python and the `pandas` library to summarize data from an Excel file and save the results in a new Excel file.

---

## Features

- **Automated Analysis**: Summarizes data from an Excel file, including:
  - First few rows of the data.
  - Descriptive statistics for numerical columns.
  - Count of unique values in each column.
  - Sum of null values in each column.
  - Mean, median, and mode for numerical and categorical columns.
- **Self-Contained**: Creates a Python virtual environment and installs all required dependencies.
- **Dynamic Script Generation**: Generates and runs a Python script dynamically to perform the analysis.
- **Structured Output**: Saves the analysis results in an Excel file in the `reports/` folder.

---

## Prerequisites

- **PowerShell**: The script is written in PowerShell and should be run in a PowerShell terminal.
- **Python**: Python 3.x must be installed on your system. The script will create a virtual environment and install the required Python packages.

---

## Directory Structure

The script creates the following directory structure:

```
.
├── src/                # Contains the input Excel file (e.g., reporte.xlsx)
├── analyzers/          # Contains the dynamically generated Python script
├── reports/            # Contains the output Excel file with analysis results
├── .venv/              # Python virtual environment
└── run.ps1             # Main PowerShell script
```

---

## Setup

1. **Place Your Excel File**:
   - Place your Excel file (e.g., `reporte.xlsx`) in the `src/` folder.

3. **Run the Script**:
   - Open a PowerShell terminal in the repository's root directory.
   - Run the script:
     ```powershell
     ./run.ps1
     ```

---

## How It Works

1. **Directory Structure**:
   - The script creates the necessary directories (`src/`, `analyzers/`, `reports/`) if they don't already exist.

2. **Virtual Environment**:
   - A Python virtual environment (`.venv/`) is created, and the required packages (`pandas`, `openpyxl`) are installed.

3. **Python Script Generation**:
   - The script dynamically generates a Python script (`analyzers/excelAnalyzer.py`) to analyze the Excel file.

4. **Analysis**:
   - The Python script reads the Excel file, performs the analysis, and saves the results in `reports/report.xlsx`.

5. **Output**:
   - The analysis results are saved in the `reports/` folder as `report.xlsx`.

---

## Example Output

The output Excel file (`reports/report.xlsx`) contains the following sheets:

- **First Few Rows**: The first 5 rows of the data.
- **Data Info**: Summary of the data (e.g., column names, data types, non-null counts).
- **Descriptive Statistics**: Descriptive statistics for numerical columns (e.g., mean, std, min, max).
- **Unique Values Count**: Count of unique values in each column.
- **Null Values Count**: Sum of null values in each column.
- **Mean of Numerical Columns**: Mean of numerical columns.
- **Median of Numerical Columns**: Median of numerical columns.
- **Mode of Categorical Columns**: Mode of categorical columns.

---

## Customization

- **Input File**: Replace `src/reporte.xlsx` with your own Excel file.
- **Output File**: Modify the `report_file` variable in the Python script to change the output file name or location.

---

## Troubleshooting

### Common Issues

1. **Python Not Found**:
   - Ensure Python is installed and added to your system's PATH.
   - You can check by running `python --version` in your terminal.

2. **Virtual Environment Activation Fails**:
   - On macOS/Linux, ensure the activation command is `. .venv/bin/activate`.
   - On Windows, ensure the activation command is `.\.venv\Scripts\Activate.ps1`.

3. **Missing Excel File**:
   - Ensure the Excel file is placed in the `src/` folder and named `reporte.xlsx`.

---

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

---