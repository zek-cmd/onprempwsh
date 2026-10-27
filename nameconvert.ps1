# Name Format Converter Script
# Converts names from "LastName, FirstName MiddleInitial" to "FirstName MiddleInitial. LastName"
# Input: CSV file with 'FullName' header
# Output: Display results (can be copied with Ctrl+C)

Add-Type -AssemblyName System.Windows.Forms

# Function to open file dialog for CSV selection
function Get-NamesFile {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $openFileDialog.Title = "Select CSV file with FullName column"
    $openFileDialog.FileName = "names.csv"
    
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        return $openFileDialog.FileName
    }
    else {
        Write-Host "File selection cancelled." -ForegroundColor Yellow
        return $null
    }
}

# Function to convert name format
function Convert-NameFormat {
    param([string]$FullName)
    
    if ([string]::IsNullOrWhiteSpace($FullName)) {
        return $FullName
    }
    
    # Trim whitespace
    $FullName = $FullName.Trim()
    
    # Check if the name contains a comma (LastName, FirstName format)
    if ($FullName -match ',') {
        # Split by comma
        $parts = $FullName -split ',', 2
        
        $lastName = $parts[0].Trim()
        $firstAndMiddle = $parts[1].Trim()
        
        # Split first name and middle initial
        $nameParts = $firstAndMiddle -split '\s+'
        
        $firstName = $nameParts[0]
        $middleInitial = ""
        
        if ($nameParts.Count -gt 1) {
            $middleInitial = $nameParts[1]
        }
        
        # Clean up middle initial (remove existing periods, then add one)
        if (-not [string]::IsNullOrWhiteSpace($middleInitial)) {
            $middleInitial = $middleInitial.Trim().TrimEnd('.')
            $middleInitial = "$middleInitial."
        }
        
        # Build the new format: FirstName MiddleInitial. LastName
        if ([string]::IsNullOrWhiteSpace($middleInitial)) {
            $converted = "$firstName $lastName"
        }
        else {
            $converted = "$firstName $middleInitial $lastName"
        }
        
        return $converted.Trim()
    }
    else {
        # No comma found, return as-is
        Write-Warning "Name '$FullName' doesn't match expected format (LastName, FirstName MI)"
        return $FullName
    }
}

# Main execution
function Main {
    Clear-Host
    
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "     NAME FORMAT CONVERTER" -ForegroundColor Cyan
    Write-Host "     LastName, FirstName MI -> FirstName MI. LastName" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Get CSV file
    Write-Host "Select the CSV file containing FullName column..." -ForegroundColor White
    $csvPath = Get-NamesFile
    
    if (-not $csvPath) {
        Write-Host "Operation cancelled. Exiting..." -ForegroundColor Red
        Write-Host "Press Enter to exit..." -ForegroundColor Gray
        Read-Host
        return
    }
    
    # Read CSV file
    try {
        $csvData = Import-Csv -Path $csvPath
        
        if (-not $csvData) {
            Write-Host "Failed to read CSV file or file is empty." -ForegroundColor Red
            Write-Host "Press Enter to exit..." -ForegroundColor Gray
            Read-Host
            return
        }
        
        # Check for FullName column
        $columns = $csvData[0].PSObject.Properties.Name
        if ($columns -notcontains 'FullName') {
            Write-Host "CSV file does not contain a 'FullName' column." -ForegroundColor Red
            Write-Host "Available columns: $($columns -join ', ')" -ForegroundColor Yellow
            Write-Host "Press Enter to exit..." -ForegroundColor Gray
            Read-Host
            return
        }
        
        Write-Host "`nFile loaded successfully: $csvPath" -ForegroundColor Green
        Write-Host "Total names found: $($csvData.Count)" -ForegroundColor Green
        Write-Host ""
        Write-Host "=" * 70 -ForegroundColor Cyan
        Write-Host "  ORIGINAL NAME                          CONVERTED NAME" -ForegroundColor Yellow
        Write-Host "=" * 70 -ForegroundColor Cyan
        
        $convertedNames = @()
        
        foreach ($row in $csvData) {
            $originalName = $row.FullName.Trim()
            
            if (-not [string]::IsNullOrWhiteSpace($originalName)) {
                $convertedName = Convert-NameFormat -FullName $originalName
                $convertedNames += $convertedName
                
                Write-Host ("  {0,-38} {1}" -f $originalName, $convertedName) -ForegroundColor White
            }
        }
        
        Write-Host "=" * 70 -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Conversion complete! $($convertedNames.Count) names processed." -ForegroundColor Green
        Write-Host ""
        Write-Host "--- Copy-Ready List (select and press Ctrl+C to copy) ---" -ForegroundColor Yellow
        Write-Host ""
        
        # Display names in a clean format for easy copying
        foreach ($name in $convertedNames) {
            Write-Host $name
        }
        
        Write-Host ""
        Write-Host "--- End of List ---" -ForegroundColor Yellow
        
    }
    catch {
        Write-Host "Error reading CSV file: $_" -ForegroundColor Red
    }
    
    Write-Host "`nPress Enter to exit..." -ForegroundColor Gray
    Read-Host
}

# Run the script
Main