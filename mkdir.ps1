# Define colors for output
$GREEN = "Green"
$YELLOW = "Yellow"
$NC = "White"

# Define the root directory (current directory)
$rootDir = "."

# Define the directory names
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