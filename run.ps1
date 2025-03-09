# Define colors for output
$GREEN = "Green"
$YELLOW = "Yellow"

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
}

function createStructure {
    Write-Host "🏗️ Creating Directory Structure" -ForegroundColor $YELLOW

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
    python -m pip install pandas openpyxl

    # Create subdirectories
    Write-Host "Creating directory structure..." -ForegroundColor $GREEN
    $directories = @(
        "src",
        "analyzers",
        "reports"
    )
    foreach ($dir in $directories) {
        if (-not (Test-Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force
        }
    }

    # Create empty Python files
    $files = @(
        "analyzers/excelAnalyzer.py"
    )
    foreach ($file in $files) {
        if (-not (Test-Path $file)) {
            New-Item -Path $file -ItemType File -Force
        }
    }
}

function generateAnalysis {
    Write-Host "🏗️ Generating Analysis" -ForegroundColor $YELLOW

    # Check if the src file exists
    $srcFile = "src/reporte.xlsx"
    if (-not (Test-Path $srcFile)) {
        Write-Host "❌ Source file '$srcFile' not found. Please place your Excel file in the 'src' folder." -ForegroundColor "Red"
        exit
    }

    # Run the Python script to analyze the Excel file
    Write-Host "Running Excel Analyzer..." -ForegroundColor $GREEN
    python analyzers/excelAnalyzer.py $srcFile "reports"

    Write-Host "✅ Analysis completed. Check the 'reports' folder for results." -ForegroundColor $GREEN
}

function analysis {
    Write-Host "🏗️ Starting Analysis Process" -ForegroundColor $YELLOW

    # Call functions to create structure and generate tables
    createStructure
    excelAnalyzer

    # Activate virtual environment
    .\.venv\Scripts\Activate.ps1

    # Generate analysis
    generateAnalysis

    Write-Host "🏗️ Analysis process completed successfully." -ForegroundColor $YELLOW
}

# Call the main function
analysis