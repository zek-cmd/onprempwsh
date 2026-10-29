# Name Format Converter Script
# Converts names from "LastName, FirstName MiddleInitial" to "FirstName MiddleInitial. LastName"
# Input: CSV file with 'FullName' header (comma or semicolon separated)
# Output: Display results (can be copied with Ctrl+C)

Add-Type -AssemblyName System.Windows.Forms

# Function to open file dialog for CSV selection
function Get-NamesFile {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $openFileDialog.Title = "Select CSV file with FullName column"
    $openFileDialog.FileName = "names.csv"
    $openFileDialog.CheckFileExists = $true

    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    }
    else {
        Write-Host "File selection cancelled." -ForegroundColor Yellow
        return $null
    }
}

# Work out whether the file uses a comma or a semicolon between columns
# (Excel in some regional settings saves with semicolons)
function Get-CsvDelimiter {
    param([string]$Path)

    $headerLine = Get-Content -Path $Path -TotalCount 1 -ErrorAction Stop
    if ($null -eq $headerLine) { return ',' }

    $commas = ([regex]::Matches($headerLine, ',')).Count
    $semicolons = ([regex]::Matches($headerLine, ';')).Count

    if ($semicolons -gt $commas) { return ';' }
    return ','
}

# Function to convert name format
# Returns an object: Name (converted text) and Note (empty when all went well)
function Convert-NameFormat {
    param([string]$FullName)

    $result = [PSCustomObject]@{ Name = ''; Note = '' }

    if ([string]::IsNullOrWhiteSpace($FullName)) {
        return $result
    }

    # Trim whitespace and collapse repeated spaces inside the name
    $FullName = ($FullName -replace '\s+', ' ').Trim()

    # Check if the name contains a comma (LastName, FirstName format)
    if ($FullName -notmatch ',') {
        # No comma: name is probably already in the new format, or is wrong
        $result.Name = $FullName
        $result.Note = "No comma found - kept as typed, please check"
        return $result
    }

    # Split only at the first comma
    $parts = $FullName -split ',', 2

    $lastName = $parts[0].Trim()
    $firstAndMiddle = $parts[1].Trim()

    if ([string]::IsNullOrWhiteSpace($lastName) -or [string]::IsNullOrWhiteSpace($firstAndMiddle)) {
        $result.Name = $FullName
        $result.Note = "Missing last name or first name - kept as typed, please check"
        return $result
    }

    # Split the first name(s) from the middle initial.
    # First names can have several words (Mary Ann, Princess Erika).
    # The middle initial is the LAST word, and only when it is a single letter (G or G.).
    $nameParts = @($firstAndMiddle -split ' ' | Where-Object { $_ -ne '' })

    $note = ""
    $middleInitial = ""

    if ($nameParts.Count -gt 1 -and $nameParts[-1] -match '^[A-Za-z]\.?$') {
        $middleInitial = $nameParts[-1].TrimEnd('.')
        $firstName = ($nameParts[0..($nameParts.Count - 2)]) -join ' '
    }
    else {
        # No single-letter middle initial at the end (e.g. "Chester John")
        $firstName = $firstAndMiddle
        $note = "No middle initial found - please check"
    }

    # Add the period after the middle initial
    if ($middleInitial -ne '') {
        $middleInitial = "$middleInitial."
    }

    # Build the new format: FirstName MiddleInitial. LastName
    if ([string]::IsNullOrWhiteSpace($middleInitial)) {
        $converted = "$firstName $lastName"
    }
    else {
        $converted = "$firstName $middleInitial $lastName"
    }

    $result.Name = $converted.Trim()
    $result.Note = $note
    return $result
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

    try {
        if (-not (Test-Path -LiteralPath $csvPath -PathType Leaf)) {
            Write-Host "The file was not found: $csvPath" -ForegroundColor Red
            Write-Host "Press Enter to exit..." -ForegroundColor Gray
            Read-Host
            return
        }

        # Read CSV file (detect comma or semicolon)
        $delimiter = Get-CsvDelimiter -Path $csvPath
        $csvData = @(Import-Csv -Path $csvPath -Delimiter $delimiter)

        if ($csvData.Count -eq 0) {
            Write-Host "Failed to read CSV file or file has no data rows." -ForegroundColor Red
            Write-Host "Press Enter to exit..." -ForegroundColor Gray
            Read-Host
            return
        }

        # Find the FullName column (ignore case, spaces, and a hidden BOM mark)
        $columns = @($csvData[0].PSObject.Properties.Name)
        $fullNameColumn = $columns | Where-Object { ($_ -replace '[\uFEFF\s]', '') -ieq 'FullName' } | Select-Object -First 1

        if (-not $fullNameColumn) {
            Write-Host "CSV file does not contain a 'FullName' column." -ForegroundColor Red
            Write-Host "Available columns: $($columns -join ', ')" -ForegroundColor Yellow
            Write-Host "Press Enter to exit..." -ForegroundColor Gray
            Read-Host
            return
        }

        Write-Host "`nFile loaded successfully: $csvPath" -ForegroundColor Green
        Write-Host "Total rows found: $($csvData.Count)" -ForegroundColor Green
        Write-Host ""
        Write-Host ("=" * 70) -ForegroundColor Cyan
        Write-Host "  ORIGINAL NAME                          CONVERTED NAME" -ForegroundColor Yellow
        Write-Host ("=" * 70) -ForegroundColor Cyan

        $convertedNames = @()
        $reviewLines = @()
        $skippedBlank = 0

        foreach ($row in $csvData) {
            $originalName = [string]$row.$fullNameColumn
            $originalName = ($originalName -replace '\s+', ' ').Trim()

            if ([string]::IsNullOrWhiteSpace($originalName)) {
                $skippedBlank++
                continue
            }

            $converted = Convert-NameFormat -FullName $originalName
            $convertedNames += $converted.Name

            if ($converted.Note) {
                $reviewLines += "  $originalName  ->  $($converted.Note)"
            }

            Write-Host ("  {0,-38} {1}" -f $originalName, $converted.Name) -ForegroundColor White
        }

        Write-Host ("=" * 70) -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Conversion complete! $($convertedNames.Count) names processed." -ForegroundColor Green
        if ($skippedBlank -gt 0) {
            Write-Host "Skipped $skippedBlank empty row(s)." -ForegroundColor Yellow
        }
        if ($reviewLines.Count -gt 0) {
            Write-Host "$($reviewLines.Count) name(s) need a check:" -ForegroundColor Yellow
            $reviewLines | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        }
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
        Write-Host "Error reading CSV file: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Check that the file is a CSV and is not open in another program." -ForegroundColor Yellow
    }

    Write-Host "`nPress Enter to exit..." -ForegroundColor Gray
    Read-Host
}

# Run the script
Main
