# ==============================
# Configuration
# ==============================

Add-Type -AssemblyName System.Windows.Forms

$TargetFolders = @("Desktop", "Downloads", "Documents")
$Report = @()

# ==============================
# Select Folder to Scan
# ==============================

Write-Host "Select the folder to scan for missing Desktop, Downloads, and Documents folders..." -ForegroundColor Cyan

$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$folderDialog.Description = "Select folder to scan for missing Desktop, Downloads, and Documents folders"
$folderDialog.RootFolder = [System.Environment+SpecialFolder]::Desktop

if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $BasePath = $folderDialog.SelectedPath
    Write-Host "Selected scan folder: $BasePath" -ForegroundColor Green
} else {
    Write-Host "No folder selected. Exiting." -ForegroundColor Red
    exit
}

# ==============================
# Recursive Scanner
# ==============================

function Scan-Folder {
    param (
        [string]$Path
    )

    # Check if this folder contains any of the target folders
    $HasTarget = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
                 Where-Object { $TargetFolders -contains $_.Name }

    if ($HasTarget) {

        # Initialize fixed schema
        $Desktop   = "Missing"
        $Downloads = "Missing"
        $Documents = "Missing"

        foreach ($Folder in $TargetFolders) {
            $TargetPath = Join-Path $Path $Folder

            if (Test-Path $TargetPath) {
                Set-Variable -Name $Folder -Value "Exists"
            }
            else {
                New-Item -ItemType Directory -Path $TargetPath | Out-Null
                Set-Variable -Name $Folder -Value "Created"
            }
        }

        # Create report row
        $Row = [PSCustomObject]@{
            ParentPath = $Path
            Desktop    = $Desktop
            Downloads  = $Downloads
            Documents  = $Documents
        }

        $script:Report += $Row
        return
    }

    # Continue recursion
    Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        Scan-Folder -Path $_.FullName
    }
}

# ==============================
# Start Scan
# ==============================

Write-Host "Scanning folders..." -ForegroundColor Cyan
Get-ChildItem -Path $BasePath -Directory | ForEach-Object {
    Scan-Folder -Path $_.FullName
}
Write-Host "Scan complete." -ForegroundColor Green

# ==============================
# Select where to export CSV
# ==============================

Write-Host "`nSelect where to save the CSV report..." -ForegroundColor Cyan

$saveDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$saveDialog.FileName = "Folder_Audit_Report.csv"
$saveDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
$saveDialog.Title = "Select where to save the CSV report"

if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $CsvPath = $saveDialog.FileName
    Write-Host "Selected CSV save location: $CsvPath" -ForegroundColor Green
} else {
    Write-Host "No save location selected. Exiting." -ForegroundColor Red
    exit
}

# ==============================
# Export to CSV
# ==============================

$Report | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding ASCII

Write-Host "`nCSV report successfully exported to:" -ForegroundColor Green
Write-Host $CsvPath -ForegroundColor Yellow