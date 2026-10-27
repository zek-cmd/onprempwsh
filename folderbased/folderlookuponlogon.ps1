# ==============================
# Configuration
# ==============================

$TargetFolders = @("Desktop", "Downloads", "Documents")
$BasePath = 'C:\$folderredirect'
$CsvPath = "C:\Users\Administrator\Desktop\scripts\csv\Folder_Audit_Report.csv"
$Report = @()

# ==============================
# Function to check and create folders
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
                try {
                    New-Item -ItemType Directory -Path $TargetPath -Force | Out-Null
                    Set-Variable -Name $Folder -Value "Created"
                }
                catch {
                    Set-Variable -Name $Folder -Value "Error: $_"
                }
            }
        }

        # Create report row
        $Row = [PSCustomObject]@{
            ParentPath = $Path
            Desktop    = $Desktop
            Downloads  = $Downloads
            Documents  = $Documents
            ScanDate   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
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

Write-Host "Starting folder scan at: $BasePath" -ForegroundColor Cyan

# Check if base path exists
if (-not (Test-Path $BasePath)) {
    Write-Host "Error: Base path '$BasePath' does not exist. Exiting." -ForegroundColor Red
    exit 1
}

Write-Host "Scanning folders..." -ForegroundColor Cyan

try {
    # Get all immediate subdirectories of the base path
    $subFolders = Get-ChildItem -Path $BasePath -Directory -ErrorAction Stop
    
    if ($subFolders.Count -eq 0) {
        Write-Host "No subfolders found in $BasePath" -ForegroundColor Yellow
        $Report = [PSCustomObject]@{
            ParentPath = $BasePath
            Desktop    = "NoSubfolders"
            Downloads  = "NoSubfolders"
            Documents  = "NoSubfolders"
            ScanDate   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    }
    else {
        # Scan each subfolder
        $subFolders | ForEach-Object {
            Scan-Folder -Path $_.FullName
        }
    }
}
catch {
    Write-Host "Error scanning folders: $_" -ForegroundColor Red
    exit 1
}

Write-Host "Scan complete. Found $($Report.Count) locations with target folders." -ForegroundColor Green

# ==============================
# Export to CSV (replace if exists)
# ==============================

try {
    # Ensure output directory exists
    $outputDir = Split-Path $CsvPath -Parent
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    
    # Export to CSV (will replace if file exists)
    $Report | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding ASCII -Force
    
    Write-Host "CSV report successfully exported to:" -ForegroundColor Green
    Write-Host $CsvPath -ForegroundColor Yellow
    Write-Host "Total records: $($Report.Count)" -ForegroundColor Cyan
}
catch {
    Write-Host "Error exporting CSV: $_" -ForegroundColor Red
    exit 1
}

# ==============================
# Optional: Display summary
# ==============================

if ($Report.Count -gt 0) {
    Write-Host "`nSummary:" -ForegroundColor Cyan
    $Report | ForEach-Object {
        Write-Host "  $($_.ParentPath):" -ForegroundColor White
        Write-Host "    Desktop: $($_.Desktop)" -ForegroundColor $(if ($_.Desktop -eq "Created") { "Green" } else { "Gray" })
        Write-Host "    Downloads: $($_.Downloads)" -ForegroundColor $(if ($_.Downloads -eq "Created") { "Green" } else { "Gray" })
        Write-Host "    Documents: $($_.Documents)" -ForegroundColor $(if ($_.Documents -eq "Created") { "Green" } else { "Gray" })
    }
}