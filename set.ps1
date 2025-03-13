# Define colors for output
$GREEN = "Green"
$YELLOW = "Yellow"

function createStructure {
    Write-Host "🏗️ Creating Directory Structure" -ForegroundColor $YELLOW

    # Define the root directory (current directory)
    $rootDir = "."

    # Create Python virtual environment
    if (-not (Test-Path ".venv")) {
        Write-Host "Creating Python virtual environment..." -ForegroundColor $GREEN
        python -m venv .venv
    }

    # Activate the virtual environment
    Write-Host "Activating virtual environment..." -ForegroundColor $GREEN
    .\.venv\Scripts\Activate.ps1

    # Upgrade pip and install required packages
    Write-Host "Installing required Python packages..." -ForegroundColor $GREEN
    python -m pip install --upgrade pip
    python -m pip install pandas openpyxl msoffcrypto-tool

    # Create subdirectories
    Write-Host "Creating directory structure..." -ForegroundColor $GREEN
    $directories = @(
        "src",
        "analyzers",
        "reports"
    )
    # Loop through the directory names and create them
    foreach ($dir in $directories) {
        $fullPath = Join-Path -Path $rootDir -ChildPath $dir

        # Check if the directory already exists
        if (-not (Test-Path -Path $fullPath -PathType Container)) {
            try {
                New-Item -ItemType Directory -Path $fullPath -Force
                Write-Host "Created directory: $fullPath" -ForegroundColor $GREEN
            }
            catch {
                Write-Host "Error creating directory $($fullPath): $($_.Exception.Message)" -ForegroundColor "Red"
            }
        }
        else {
            Write-Host "Directory already exists: $fullPath" -ForegroundColor $YELLOW
        }
    }

    Write-Host "Directory structure creation complete." -ForegroundColor $GREEN

    # Create empty Python files
    $files = @(
        "analyzers/excelAnalyzer.py",
        "analyzers/passKey.py",
        "analyzers/checkCols.py"
    )
    foreach ($file in $files) {
        if (-not (Test-Path $file)) {
            New-Item -Path $file -ItemType File -Force
        }
    }
}
function excelAnalyzer {
    Write-Host "🏗️ Creating Excel Analyzer Script" -ForegroundColor $YELLOW
    # Create the Python script dynamically
    Set-Content -Path "analyzers/excelAnalyzer.py" -Value @"
import pandas as pd
import os

def summarize_excel(file_path, report_folder):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Create a dictionary to store all summary results
    summary_results = {}

    # First few rows of the data
    summary_results["First Few Rows"] = df.head()

    # Basic summary of the data
    summary_results["Data Info"] = df.info()

    # Descriptive statistics for numerical columns
    summary_results["Descriptive Statistics"] = df.describe()

    # Count of unique values in each column
    summary_results["Unique Values Count"] = df.nunique()

    # Sum of null values in each column
    summary_results["Null Values Count"] = df.isnull().sum()

    # Additional summaries (e.g., mean, median, mode) for specific columns
    if len(df.select_dtypes(include=['number']).columns) > 0:
        summary_results["Mean of Numerical Columns"] = df.mean(numeric_only=True)
        summary_results["Median of Numerical Columns"] = df.median(numeric_only=True)

    if len(df.select_dtypes(include=['object']).columns) > 0:
        summary_results["Mode of Categorical Columns"] = df.mode().iloc[0]

    # Save the results to an Excel file
    report_file = os.path.join(report_folder, "report.xlsx")
    with pd.ExcelWriter(report_file, engine='openpyxl') as writer:
        for sheet_name, data in summary_results.items():
            if isinstance(data, pd.DataFrame):
                data.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                pd.DataFrame([data]).to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"✅ Analysis report saved to: {report_file}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage: python excel_analyzer.py <src/reporte.xlsx> <reports>")
    else:
        file_path = sys.argv[1]
        report_folder = sys.argv[2]
        summarize_excel(file_path, report_folder)
"@

    Write-Host "Excel Analyzer script created successfully." -ForegroundColor $GREEN

    # Create the passKey.py script
    Set-Content -Path "analyzers/passKey.py" -Value @"
import msoffcrypto

def remove_excel_password(input_file, output_file=None):
    
    if output_file is None:
        output_file = input_file  # Overwrite the input file if no output file is specified

    # Prompt for the password
    password = input(f"Enter password for '{input_file}': ")

    try:
        # Decrypt the file using msoffcrypto-tool
        decrypted_file = output_file
        with open(input_file, "rb") as file:
            office_file = msoffcrypto.OfficeFile(file)
            office_file.load_key(password=password)  # Load the password
            with open(decrypted_file, "wb") as decrypted:
                office_file.decrypt(decrypted)  # Decrypt and save the file

        print(f"Password removed successfully. File saved to '{output_file}'.")
        return True
    except msoffcrypto.exceptions.InvalidKeyError:
        print("Incorrect password. Please try again.")
        return False
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return False


if __name__ == "__main__":
    # Example usage
    input_excel_file = "src/name.xlsx"
    output_excel_file = "src/name.xlsx"

    if remove_excel_password(input_excel_file, output_excel_file):
        print("Process completed successfully.")
    else:
        print("Process failed.")
"@

    Write-Host "PassKey script created successfully." -ForegroundColor $GREEN

    # Create the checkCols.py script
    Set-Content -Path "analyzers/checkCols.py" -Value @"
import pandas as pd

df = pd.read_excel('src/name.xlsx')
print(df.columns)
"@

    Write-Host "CheckCols script created successfully." -ForegroundColor $GREEN
}

function setProject{
    Write-Host "🏗️ Setting project" -ForegroundColor $YELLOW

    # Call functions to create structure and generate tables
    createStructure
    excelAnalyzer

    # Activate virtual environment
    .\.venv\Scripts\Activate.ps1

    Write-Host "🏗️ Set process completed successfully." -ForegroundColor $YELLOW
}
# Call the main function
setProject