# PowerShell Script: Extract Documents from User Folders Based on CSV
# Description: Finds folders matching usernames, extracts .xlsx and .docx files from them
# Usage: Run the script and select the CSV file when prompted

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

# Display starting message
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host " DOCUMENT EXTRACTION SCRIPT" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
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

# Step 4: Create main output folder
if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    Write-Host "Created output folder: $OutputFolder" -ForegroundColor Green
} else {
    Write-Host "Output folder already exists: $OutputFolder" -ForegroundColor Yellow
}
Write-Host ""

# Step 5: Find all user folders (excluding Extracted_Files)
Write-Host "Scanning for user folders..." -ForegroundColor Yellow

# Get all folders recursively, excluding the Extracted_Files folder
$AllFolders = Get-ChildItem -Path $CSVDirectory -Directory -Recurse | Where-Object {
    $_.FullName -notlike "$OutputFolder\*"
}

Write-Host "Found $($AllFolders.Count) total folders to scan." -ForegroundColor Green
Write-Host ""

# Step 6: Initialize counters
$TotalFilesCopied = 0
$UsersWithFiles = 0
$UsersWithoutFiles = 0
$UsersWithNoFolder = 0
$UserStatus = @()
$ProcessedUsernames = @()

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
    
    # Sanitize folder name (remove invalid characters for Windows folder names)
    $InvalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $SanitizedFullName = $FullName
    foreach ($Char in $InvalidChars) {
        $SanitizedFullName = $SanitizedFullName.Replace($Char, '_')
    }
    
    # Create user folder in Extracted_Files
    $UserFolder = Join-Path -Path $OutputFolder -ChildPath $SanitizedFullName
    
    # Always create user folder
    if (-not (Test-Path -Path $UserFolder)) {
        New-Item -Path $UserFolder -ItemType Directory -Force | Out-Null
    }
    
    # Find all folders matching the username (exact match or containing username)
    $MatchingFolders = $AllFolders | Where-Object { $_.Name -match [regex]::Escape($Username) }
    
    if ($MatchingFolders.Count -gt 0) {
        Write-Host "  Found $($MatchingFolders.Count) folder(s) matching username '$Username'" -ForegroundColor Green
        
        $FolderFileCount = 0
        $FileCopiedCount = 0
        
        foreach ($Folder in $MatchingFolders) {
            # Find all .xlsx and .docx files inside this folder and its subfolders
            $Files = Get-ChildItem -Path $Folder.FullName -File -Include "*.xlsx", "*.docx" -Recurse
            
            if ($Files.Count -gt 0) {
                $FolderFileCount += $Files.Count
                Write-Host "    Found $($Files.Count) files in folder: $($Folder.Name)" -ForegroundColor Gray
                
                foreach ($File in $Files) {
                    $DestinationPath = Join-Path -Path $UserFolder -ChildPath $File.Name
                    $DestinationFileName = $File.Name
                    $Counter = 1
                    
                    # Handle duplicate filenames
                    while (Test-Path -Path $DestinationPath) {
                        $NameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
                        $Extension = [System.IO.Path]::GetExtension($File.Name)
                        $DestinationFileName = "$NameWithoutExt`_$Counter$Extension"
                        $DestinationPath = Join-Path -Path $UserFolder -ChildPath $DestinationFileName
                        $Counter++
                    }
                    
                    # Copy the file
                    try {
                        Copy-Item -Path $File.FullName -Destination $DestinationPath -Force
                        Write-Host "      Copied: $($File.Name) -> $DestinationFileName" -ForegroundColor Gray
                        $FileCopiedCount++
                        $TotalFilesCopied++
                    } catch {
                        Write-Host "      ERROR copying file: $($File.Name) - $_" -ForegroundColor Red
                    }
                }
            } else {
                Write-Host "    No .xlsx or .docx files found in folder: $($Folder.Name)" -ForegroundColor Gray
            }
        }
        
        if ($FileCopiedCount -gt 0) {
            $UsersWithFiles++
            Write-Host "  Extracted $FileCopiedCount file(s) from $($MatchingFolders.Count) folder(s)" -ForegroundColor Green
            Write-Host "  Destination: $UserFolder" -ForegroundColor Green
            $UserStatus += "$FullName - $FileCopiedCount file(s) extracted from $($MatchingFolders.Count) folder(s)"
        } else {
            $UsersWithoutFiles++
            Write-Host "  No .xlsx or .docx files found in any matching folders" -ForegroundColor Yellow
            Write-Host "  Created empty folder: $UserFolder" -ForegroundColor Gray
            $UserStatus += "$FullName - No files found (empty folder created)"
        }
        
    } else {
        $UsersWithNoFolder++
        Write-Host "  No folder found matching username '$Username'" -ForegroundColor Yellow
        Write-Host "  Created empty folder: $UserFolder" -ForegroundColor Gray
        $UserStatus += "$FullName - No folder found (empty folder created)"
    }
    
    Write-Host ""
}

# Step 8: Final Summary
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host " EXTRACTION SUMMARY" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "CSV File used: $CSVPath" -ForegroundColor White
Write-Host "Total users processed: $($Users.Count)" -ForegroundColor White
Write-Host "  - Users with files extracted: $UsersWithFiles" -ForegroundColor Green
Write-Host "  - Users with folder but no files: $UsersWithoutFiles" -ForegroundColor Yellow
Write-Host "  - Users with no matching folder: $UsersWithNoFolder" -ForegroundColor Yellow
Write-Host ""
Write-Host "Total files copied: $TotalFilesCopied" -ForegroundColor Green
Write-Host ""
Write-Host "Output folder location: $OutputFolder" -ForegroundColor White
Write-Host ""
Write-Host "DETAILED USER STATUS:" -ForegroundColor Cyan
foreach ($Status in $UserStatus) {
    Write-Host "  $Status" -ForegroundColor Gray
}
Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host " EXTRACTION COMPLETED SUCCESSFULLY!" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Cyan

Pause