# Define colors for output
$GREEN = "Green"
$YELLOW = "Yellow"

function generateAnalysis {
    Write-Host "🏗️ Generating Analysis" -ForegroundColor $YELLOW

    # Check if the src file exists
    $srcFile = "src/dataHistoricaPBI.xlsx"
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

    # Generate analysis
    generateAnalysis

    Write-Host "🏗️ Analysis process completed successfully." -ForegroundColor $YELLOW
}

# Call the main function
analysis