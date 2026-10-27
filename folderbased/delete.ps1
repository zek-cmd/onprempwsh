# PowerShell Script: Delete Extracted Document Files
# Description: Deletes .xlsx and .docx files from user folders that match usernames in CSV
# Usage: Run the script and select the CSV file when prompted
# WARNING: This script permanently deletes files. Use with caution!

# Define script location as root directory
$ScriptRoot = $PSScriptRoot

# If script is run from console without a path, use current directory
if ([string]::IsNullOrEmpty($ScriptRoot)) {
    $ScriptRoot = Get-Location
}

# Function to open file dialog and select CSV file
function Select-CSVFile {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Select CSV File"
    $OpenFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $OpenFileDialog.FilterIndex = 1
    $OpenFileDialog.RestoreDirectory = $true
    $OpenFileDialog.InitialDirectory = $ScriptRoot
    
    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $OpenFileDialog.FileName
    } else {
        return $null
    }
}

# Display warning message
Write-Host "=========================================" -ForegroundColor Red
Write-Host " WARNING: FILE DELETION SCRIPT" -ForegroundColor Red
Write-Host "=========================================" -ForegroundColor Red
Write-Host "This script will PERMANENTLY DELETE .xlsx and .docx" -ForegroundColor Yellow
Write-Host "files from folders matching usernames in the CSV file." -ForegroundColor Yellow
Write-Host ""
Write-Host "Files will be DELETED from their SOURCE locations." -ForegroundColor Red
Write-Host "Files in Extracted_Files folder will NOT be affected." -ForegroundColor Green
Write-Host ""
Write-Host "Script Location: $ScriptRoot" -ForegroundColor Yellow
Write-Host ""

# Step 1: Select CSV file
Write-Host "Please select the CSV file..." -ForegroundColor Yellow

$CSVPath = Select-CSVFile

if (-not $CSVPath) {
    Write-Host "No CSV file selected. Exiting..." -ForegroundColor Red
    Pause
    Exit 0
}

Write-Host "Selected CSV file: $CSVPath" -ForegroundColor Green
Write-Host ""

# Step 2: Validate CSV file exists
if (-not (Test-Path -Path $CSVPath)) {
    Write-Host "ERROR: CSV file not found at: $CSVPath" -ForegroundColor Red
    Pause
    Exit 1
}

# Define output folder (in the same location as the CSV file)
$CSVDirectory = Split-Path -Path $CSVPath -Parent
$OutputFolder = Join-Path -Path $CSVDirectory -ChildPath "Extracted_Files"

# Step 3: Import CSV data
Write-Host "Loading CSV file..." -ForegroundColor Yellow
try {
    $Users = Import-Csv -Path $CSVPath
    Write-Host "Successfully loaded $($Users.Count) user records." -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "ERROR: Failed to import CSV file: $_" -ForegroundColor Red
    Pause
    Exit 1
}

# Step 4: Find all user folders (excluding Extracted_Files)
Write-Host "Scanning for user folders..." -ForegroundColor Yellow

# Get all folders recursively, excluding the Extracted_Files folder
$AllFolders = Get-ChildItem -Path $CSVDirectory -Directory -Recurse | Where-Object {
    $_.FullName -notlike "$OutputFolder\*"
}

Write-Host "Found $($AllFolders.Count) total folders to scan." -ForegroundColor Green
Write-Host ""

# Step 5: Confirm deletion
Write-Host "=========================================" -ForegroundColor Red
Write-Host " CONFIRM DELETION" -ForegroundColor Red
Write-Host "=========================================" -ForegroundColor Red
Write-Host ""
Write-Host "CSV File: $CSVPath" -ForegroundColor White
Write-Host "This will delete .xlsx and .docx files from folders" -ForegroundColor Yellow
Write-Host "matching usernames in the CSV file." -ForegroundColor Yellow
Write-Host ""
Write-Host "Number of users to process: $($Users.Count)" -ForegroundColor White
Write-Host ""
$Confirm = Read-Host "Type 'DELETE' to confirm and proceed"

if ($Confirm -ne "DELETE") {
    Write-Host ""
    Write-Host "Deletion cancelled. No files were deleted." -ForegroundColor Green
    Pause
    Exit 0
}

Write-Host ""
Write-Host "Confirmation received. Proceeding with deletion..." -ForegroundColor Green
Write-Host ""

# Step 6: Initialize counters
$TotalFilesDeleted = 0
$TotalFilesFailed = 0
$UsersWithFiles = 0
$UsersWithNoFiles = 0
$UsersWithNoFolder = 0
$DeletedFilesList = @()
$FailedFilesList = @()

# Step 7: Process each user
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host " PROCESSING USERS" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

foreach ($User in $Users) {
    $FullName = $User.FullName
    $Username = $User.Username
    
    # Skip if FullName or Username is empty
    if ([string]::IsNullOrWhiteSpace($FullName) -or [string]::IsNullOrWhiteSpace($Username)) {
        Write-Host "WARNING: Skipping user with missing data (FullName or Username)" -ForegroundColor Red
        continue
    }
    
    Write-Host "Processing: $FullName ($Username)" -ForegroundColor Cyan
    
    # Find all folders matching the username (exact match or containing username)
    $MatchingFolders = $AllFolders | Where-Object { $_.Name -match [regex]::Escape($Username) }
    
    if ($MatchingFolders.Count -gt 0) {
        Write-Host "  Found $($MatchingFolders.Count) folder(s) matching username '$Username'" -ForegroundColor Green
        
        $FolderFileCount = 0
        $FileDeletedCount = 0
        
        foreach ($Folder in $MatchingFolders) {
            # Find all .xlsx and .docx files inside this folder and its subfolders
            $Files = Get-ChildItem -Path $Folder.FullName -File -Include "*.xlsx", "*.docx" -Recurse
            
            if ($Files.Count -gt 0) {
                $FolderFileCount += $Files.Count
                Write-Host "    Found $($Files.Count) files in folder: $($Folder.Name)" -ForegroundColor Gray
                
                foreach ($File in $Files) {
                    try {
                        # Delete the file
                        Remove-Item -Path $File.FullName -Force
                        Write-Host "      Deleted: $($File.FullName)" -ForegroundColor Gray
                        $FileDeletedCount++
                        $TotalFilesDeleted++
                        $DeletedFilesList += $File.FullName
                    } catch {
                        Write-Host "      ERROR deleting file: $($File.FullName) - $_" -ForegroundColor Red
                        $TotalFilesFailed++
                        $FailedFilesList += "$($File.FullName) - $_"
                    }
                }
            } else {
                Write-Host "    No .xlsx or .docx files found in folder: $($Folder.Name)" -ForegroundColor Gray
            }
        }
        
        if ($FileDeletedCount -gt 0) {
            $UsersWithFiles++
            Write-Host "  Deleted $FileDeletedCount file(s) from $($MatchingFolders.Count) folder(s)" -ForegroundColor Green
        } else {
            $UsersWithNoFiles++
            Write-Host "  No .xlsx or .docx files found in any matching folders" -ForegroundColor Yellow
        }
        
    } else {
        $UsersWithNoFolder++
        Write-Host "  No folder found matching username '$Username'" -ForegroundColor Yellow
    }
    
    Write-Host ""
}

# Step 8: Final Summary
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host " DELETION SUMMARY" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "CSV File used: $CSVPath" -ForegroundColor White
Write-Host "Total users processed: $($Users.Count)" -ForegroundColor White
Write-Host "  - Users with files deleted: $UsersWithFiles" -ForegroundColor Green
Write-Host "  - Users with folder but no files: $UsersWithNoFiles" -ForegroundColor Yellow
Write-Host "  - Users with no matching folder: $UsersWithNoFolder" -ForegroundColor Yellow
Write-Host ""
Write-Host "Total files deleted: $TotalFilesDeleted" -ForegroundColor Green
if ($TotalFilesFailed -gt 0) {
    Write-Host "Total files failed to delete: $TotalFilesFailed" -ForegroundColor Red
}
Write-Host ""

if ($DeletedFilesList.Count -gt 0) {
    Write-Host "DELETED FILES:" -ForegroundColor Cyan
    foreach ($File in $DeletedFilesList) {
        Write-Host "  $File" -ForegroundColor Gray
    }
    Write-Host ""
}

if ($FailedFilesList.Count -gt 0) {
    Write-Host "FAILED TO DELETE:" -ForegroundColor Red
    foreach ($File in $FailedFilesList) {
        Write-Host "  $File" -ForegroundColor Red
    }
    Write-Host ""
}

Write-Host "=========================================" -ForegroundColor Cyan
if ($TotalFilesFailed -eq 0) {
    Write-Host " DELETION COMPLETED SUCCESSFULLY!" -ForegroundColor Green
} else {
    Write-Host " DELETION COMPLETED WITH ERRORS!" -ForegroundColor Red
}
Write-Host "=========================================" -ForegroundColor Cyan

Pause