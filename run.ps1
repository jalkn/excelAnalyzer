# excelAnalizer.ps1
param (
    [string]$filePath
)

# Check if the file path is provided
if (-not $filePath) {
    Write-Host "Usage: .\excelAnalizer.ps1 <input/reporte.xlsx>"
    exit
}

# Check if Python is installed
$pythonPath = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $pythonPath) {
    Write-Host "Python is not installed or not in the PATH. Please install Python and try again."
    exit
}

# Check if the Excel file exists
if (-not (Test-Path $filePath)) {
    Write-Host "The file '$filePath' does not exist."
    exit
}

# Define the Python script as a here-string
$pythonScript = @"
import pandas as pd

def summarize_excel(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Display the first few rows of the dataframe
    print("First few rows of the data:")
    print(df.head())

    # Basic summary of the data
    print("\nSummary of the data:")
    print(df.info())

    # Descriptive statistics for numerical columns
    print("\nDescriptive statistics for numerical columns:")
    print(df.describe())

    # Count of unique values in each column
    print("\nCount of unique values in each column:")
    print(df.nunique())

    # Sum of null values in each column
    print("\nSum of null values in each column:")
    print(df.isnull().sum())

    # Additional summaries (e.g., mean, median, mode) for specific columns
    if len(df.select_dtypes(include=['number']).columns) > 0:
        print("\nMean of numerical columns:")
        print(df.mean(numeric_only=True))

        print("\nMedian of numerical columns:")
        print(df.median(numeric_only=True))

    if len(df.select_dtypes(include=['object']).columns) > 0:
        print("\nMode of categorical columns:")
        print(df.mode().iloc[0])

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Usage: python excel_analyzer.py <path_to_excel_file>")
    else:
        file_path = sys.argv[1]
        summarize_excel(file_path)
"@

# Save the Python script to a temporary file
$tempPythonFile = "$env:TEMP\temp_excel_analyzer.py"
$pythonScript | Out-File -FilePath $tempPythonFile -Encoding utf8

# Run the Python script with the provided Excel file path
Write-Host "Analyzing Excel file: $filePath"
python $tempPythonFile $filePath

# Clean up the temporary Python file
Remove-Item -Path $tempPythonFile -Force