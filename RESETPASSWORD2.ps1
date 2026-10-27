# Active Directory Password Reset Script (CSV-Based)
# Reset passwords for users listed in names.csv (FullName + StudNo columns)
# Uses StudNo as the username for AD lookup
# Script compatibility verified for PowerShell 5.1+

# Import required assemblies for dialog boxes
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

# Configuration
$NewPassword = "password"

# Function to get domain configuration from environment variables
function Get-CurrentDomainInfo {
    [CmdletBinding()]
    param()
    
    $domainInfo = @{
        Domain = ""
        DomainDC = ""
        DomainTLD = ""
        DomainDN = ""
        NetBIOSName = ""
    }
    
    $domainInfo.Domain = $env:USERDNSDOMAIN
    
    if ([string]::IsNullOrEmpty($domainInfo.Domain)) {
        Write-Error "Cannot determine domain. USERDNSDOMAIN environment variable is not set."
        Write-Host "Make sure you are running this script from a domain-joined computer and logged in with a domain account." -ForegroundColor Red
        return $null
    }
    
    $domainParts = $domainInfo.Domain.Split('.')
    
    if ($domainParts.Count -ge 1) {
        $domainInfo.DomainDC = $domainParts[0]
    }
    if ($domainParts.Count -ge 2) {
        $domainInfo.DomainTLD = $domainParts[1]
    }
    
    $domainPartsForDN = @()
    foreach ($part in $domainParts) {
        if (-not [string]::IsNullOrEmpty($part)) {
            $domainPartsForDN += "DC=$part"
        }
    }
    $domainInfo.DomainDN = $domainPartsForDN -join ','
    
    $domainInfo.NetBIOSName = $env:USERDOMAIN
    
    return $domainInfo
}

$domainConfig = Get-CurrentDomainInfo

if (-not $domainConfig) {
    Write-Host "Press Enter to exit..." -ForegroundColor Gray
    Read-Host
    exit 1
}

$Domain = $domainConfig.Domain
$DomainDC = $domainConfig.DomainDC
$DomainTLD = $domainConfig.DomainTLD
$DomainDN = $domainConfig.DomainDN

Write-Host "`n=== DETECTED DOMAIN CONFIGURATION ===" -ForegroundColor Cyan
Write-Host "Domain: $Domain" -ForegroundColor Green
Write-Host "Domain DN: $DomainDN" -ForegroundColor Green
if (-not [string]::IsNullOrEmpty($domainConfig.NetBIOSName)) {
    Write-Host "NetBIOS Name: $($domainConfig.NetBIOSName)" -ForegroundColor Green
}
Write-Host "=" * 45 -ForegroundColor Cyan

# Function to open file dialog for names.csv
function Get-NamesFile {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $openFileDialog.Title = "Select names.csv file"
    $openFileDialog.FileName = "names.csv"
    
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        return $openFileDialog.FileName
    }
    else {
        Write-Host "File selection cancelled." -ForegroundColor Yellow
        return $null
    }
}

# Function to open Save As dialog for results CSV
function Get-ResultsSavePath {
    param(
        [string]$DefaultFileName = "AD_Password_Reset_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    )
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveFileDialog.Title = "Save Password Reset Results CSV Report"
    $saveFileDialog.FileName = $DefaultFileName
    $saveFileDialog.OverwritePrompt = $true
    $saveFileDialog.CreatePrompt = $false
    $saveFileDialog.RestoreDirectory = $true
    
    if ($saveFileDialog.ShowDialog() -eq 'OK') {
        return $saveFileDialog.FileName
    }
    else {
        Write-Host "Save cancelled by user." -ForegroundColor Yellow
        return $null
    }
}

# Function to read entries from CSV file (FullName + StudNo columns)
function Read-NamesFromCSV {
    param([string]$FilePath)
    
    if (-not (Test-Path -Path $FilePath -PathType Leaf)) {
        Write-Error "File not found: $FilePath"
        return @()
    }
    
    Write-Host "Reading entries from CSV file..." -ForegroundColor White
    
    try {
        $csvData = Import-Csv -Path $FilePath
        
        if (-not $csvData) {
            Write-Error "Failed to read CSV file or file is empty: $FilePath"
            return @()
        }
        
        $columns = $csvData[0].PSObject.Properties.Name
        
        if ($columns -notcontains 'FullName') {
            Write-Error "CSV file does not contain a 'FullName' column."
            Write-Host "Available columns: $($columns -join ', ')" -ForegroundColor Yellow
            return @()
        }
        
        if ($columns -notcontains 'StudNo') {
            Write-Error "CSV file does not contain a 'StudNo' column."
            Write-Host "Available columns: $($columns -join ', ')" -ForegroundColor Yellow
            return @()
        }
        
        $entries = @()
        foreach ($row in $csvData) {
            $fullName = $row.FullName.Trim()
            $studNo = $row.StudNo.Trim()
            
            if (-not [string]::IsNullOrWhiteSpace($studNo)) {
                $entries += [PSCustomObject]@{
                    FullName = $fullName
                    StudNo   = $studNo
                }
            }
            else {
                Write-Warning "Row has empty StudNo. Skipping."
            }
        }
        
        Write-Host "Successfully read $($entries.Count) entries from CSV file." -ForegroundColor Green
        return $entries
    }
    catch {
        Write-Error "Error reading CSV file: $_"
        return @()
    }
}

# Function to reset password for an AD user (looked up by StudNo as username)
function Reset-ADUserPassword {
    param(
        [string]$FullName,
        [string]$StudNo,
        [string]$Password
    )
    
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        Write-Error "Active Directory module is not available. Please install RSAT tools."
        return @{
            Status = "Failed"
            Message = "AD module not available"
            Username = ""
        }
    }
    
    $Username = $StudNo.Trim()
    
    if ([string]::IsNullOrWhiteSpace($Username)) {
        Write-Error "Student number is empty for: $FullName"
        return @{
            Status = "Failed"
            Message = "Student number is empty"
            Username = ""
        }
    }
    
    Write-Host "`nLooking up user: $FullName" -ForegroundColor White
    Write-Host "  StudNo / Username: $Username" -ForegroundColor Gray
    
    $adUser = $null
    try {
        $adUser = Get-ADUser -Identity $Username -Properties DistinguishedName, Name, SamAccountName, PasswordLastSet, Enabled
    }
    catch {
        # User not found by StudNo
    }
    
    if (-not $adUser) {
        Write-Host "  ERROR: User with StudNo '$Username' not found in Active Directory." -ForegroundColor Red
        return @{
            Status = "Failed"
            Message = "User not found in AD"
            Username = $Username
        }
    }
    
    Write-Host "  Found user: $($adUser.Name)" -ForegroundColor Green
    Write-Host "  Username: $($adUser.SamAccountName)" -ForegroundColor Gray
    Write-Host "  Account Enabled: $($adUser.Enabled)" -ForegroundColor Gray
    
    if ($adUser.PasswordLastSet) {
        Write-Host "  Password Last Set: $($adUser.PasswordLastSet)" -ForegroundColor Gray
    }
    else {
        Write-Host "  Password Last Set: Never" -ForegroundColor Gray
    }
    
    if (-not $adUser.Enabled) {
        Write-Host "  WARNING: Account is disabled. Password will still be reset." -ForegroundColor Yellow
    }
    
    try {
        $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        Set-ADAccountPassword -Identity $adUser.DistinguishedName -NewPassword $securePassword -Reset
        
        Write-Host "  SUCCESS: Password reset to '$Password'." -ForegroundColor Green
        
        return @{
            Status = "Success"
            Message = "Password reset successfully"
            Username = $adUser.SamAccountName
            DisplayName = $adUser.Name
        }
    }
    catch {
        Write-Host "  ERROR: Failed to reset password: $_" -ForegroundColor Red
        
        return @{
            Status = "Failed"
            Message = "Password reset failed: $_"
            Username = $adUser.SamAccountName
            DisplayName = $adUser.Name
        }
    }
}

# Function to display summary before execution
function Show-Summary {
    param(
        [array]$Entries,
        [string]$Password,
        [string]$NamesFile
    )
    
    Write-Host "`n=== PASSWORD RESET SUMMARY ===" -ForegroundColor Cyan
    Write-Host "Names file: $NamesFile" -ForegroundColor Yellow
    Write-Host "Number of users to process: $($Entries.Count)" -ForegroundColor Yellow
    Write-Host "New Password: $Password" -ForegroundColor Yellow
    Write-Host "Domain: $Domain" -ForegroundColor Yellow
    Write-Host "Lookup Method: StudNo as username" -ForegroundColor Yellow
    
    Write-Host "`n--- FullName -> StudNo ---" -ForegroundColor White
    
    $maxNameLength = 0
    foreach ($entry in $Entries) {
        if ($entry.FullName.Length -gt $maxNameLength) {
            $maxNameLength = $entry.FullName.Length
        }
    }
    $maxNameLength += 5
    
    foreach ($entry in $Entries) {
        $paddedName = $entry.FullName.PadRight($maxNameLength)
        Write-Host "$paddedName -> $($entry.StudNo)" -ForegroundColor Gray
    }
    
    Write-Host "`n=================================" -ForegroundColor Cyan
}

# Main execution
function Main {
    Clear-Host
    
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "     ACTIVE DIRECTORY PASSWORD RESET TOOL" -ForegroundColor Cyan
    Write-Host "     (CSV-Based: FullName + StudNo columns)" -ForegroundColor Cyan
    Write-Host "     (Lookup by StudNo as username)" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Checking prerequisites..." -ForegroundColor White
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        Write-Host "WARNING: Active Directory module not found." -ForegroundColor Red
        Write-Host "Please install Remote Server Administration Tools (RSAT)." -ForegroundColor Red
        Write-Host "Script will continue but cannot reset passwords without AD module." -ForegroundColor Yellow
        $continue = Read-Host "Continue anyway? (yes/no)"
        if ($continue -ne 'yes') {
            Write-Host "Script terminated by user." -ForegroundColor Yellow
            return
        }
    }
    
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    
    # Step 1: Get names CSV file
    Write-Host "`nStep 1: Select names.csv file" -ForegroundColor White
    $namesFile = Get-NamesFile
    if (-not $namesFile) {
        Write-Host "Operation cancelled by user." -ForegroundColor Red
        return
    }
    
    # Step 2: Read entries from CSV file (FullName + StudNo)
    $entries = Read-NamesFromCSV -FilePath $namesFile
    if ($entries.Count -eq 0) {
        Write-Host "No valid entries found in the file." -ForegroundColor Red
        Write-Host "Make sure your CSV has both 'FullName' and 'StudNo' column headers." -ForegroundColor Yellow
        return
    }
    
    # Step 3: Confirm or change password
    Write-Host "`nStep 3: Password configuration" -ForegroundColor White
    Write-Host "Default password: $NewPassword" -ForegroundColor Yellow
    $useCustomPassword = Read-Host "Do you want to use a different password? (Y/N)"
    
    $resetPassword = $NewPassword
    if ($useCustomPassword -eq 'Y' -or $useCustomPassword -eq 'y') {
        $customPassword = Read-Host "Enter the new password"
        if (-not [string]::IsNullOrWhiteSpace($customPassword)) {
            $resetPassword = $customPassword
        }
    }
    
    # Step 4: Show summary
    Show-Summary -Entries $entries -Password $resetPassword -NamesFile $namesFile
    
    # Step 5: Ask for confirmation
    Write-Host "`n=== CONFIRMATION REQUIRED ===" -ForegroundColor Red
    Write-Host "This will reset passwords for $($entries.Count) users to: '$resetPassword'" -ForegroundColor Yellow
    Write-Host "  Domain: $Domain" -ForegroundColor Yellow
    Write-Host "  Lookup: By StudNo as username" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "WARNING: This will forcefully change user passwords." -ForegroundColor Red
    Write-Host "Users will not be able to log in with their old password." -ForegroundColor Red
    Write-Host "Consider whether you want to force password change at next logon." -ForegroundColor Yellow
    
    $forceChangeChoice = Read-Host "`nForce users to change password at next logon? (Y/N)"
    $forceChange = ($forceChangeChoice -eq 'Y' -or $forceChangeChoice -eq 'y')
    
    $confirmation = Read-Host "`nType 'YES' to proceed, anything else to cancel"
    
    if ($confirmation -ne "YES") {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        return
    }
    
    # Step 6: Reset passwords
    Write-Host "`nStarting password resets..." -ForegroundColor Green
    
    $results = @()
    $successCount = 0
    $failCount = 0
    
    foreach ($entry in $entries) {
        Write-Host "`n" + ("="*50) -ForegroundColor DarkGray
        Write-Host "Processing: $($entry.FullName) (StudNo: $($entry.StudNo))" -ForegroundColor White
        
        $result = Reset-ADUserPassword -FullName $entry.FullName -StudNo $entry.StudNo -Password $resetPassword
        
        # Apply force password change if selected
        if ($result.Status -eq "Success" -and $forceChange) {
            try {
                $adUser = Get-ADUser -Identity $entry.StudNo
                Set-ADUser -Identity $adUser.DistinguishedName -ChangePasswordAtLogon $true
                Write-Host "  NOTE: User will be required to change password at next logon." -ForegroundColor Yellow
                $result.Message += " (Change at logon enforced)"
            }
            catch {
                Write-Host "  WARNING: Could not set 'change password at logon' flag: $_" -ForegroundColor Yellow
            }
        }
        
        $resultObject = [PSCustomObject]@{
            FullName = $entry.FullName
            StudNo   = $entry.StudNo
            Username = if ($result.Username) { $result.Username } else { "" }
            Status   = if ($result.Status) { $result.Status } else { "Failed" }
            Message  = if ($result.Message) { $result.Message } else { "Unknown error" }
            NewPassword = $resetPassword
            ForceChangeAtLogon = $forceChange
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        $results += $resultObject
        
        if ($result.Status -eq "Success") {
            $successCount++
        }
        else {
            $failCount++
        }
        
        Start-Sleep -Milliseconds 100
    }
    
    # Step 7: Display final summary
    Write-Host "`n" + ("="*60) -ForegroundColor Cyan
    Write-Host "PASSWORD RESET OPERATION COMPLETE" -ForegroundColor Cyan
    Write-Host ("="*60) -ForegroundColor Cyan
    
    if ($successCount -gt 0) {
        Write-Host "Successfully reset: $successCount passwords" -ForegroundColor Green
    }
    
    if ($failCount -gt 0) {
        Write-Host "Failed: $failCount users" -ForegroundColor Red
    }
    
    if ($results.Count -gt 0) {
        Write-Host "`n--- Detailed Results ---" -ForegroundColor White
        $results | Format-Table FullName, StudNo, Status -AutoSize | Out-Host
    }
    
    # Step 8: Let user choose where to save the results CSV
    Write-Host "`n" + ("-"*60) -ForegroundColor DarkGray
    Write-Host "Step 8: Save Results Report" -ForegroundColor White
    
    # Generate default filename with timestamp
    $defaultFileName = "AD_Password_Reset_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    # Show Save As dialog
    $exportPath = Get-ResultsSavePath -DefaultFileName $defaultFileName
    
    if ($exportPath) {
        # Export results to the selected location
        $results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nResults exported to: $exportPath" -ForegroundColor Green
        
        # Step 9: Offer to open the results file
        $openResults = Read-Host "`nOpen the results CSV file? (Y/N)"
        if ($openResults -eq 'y' -or $openResults -eq 'Y') {
            Invoke-Item $exportPath
        }
    } else {
        Write-Host "`nResults not exported. Data is still available in memory." -ForegroundColor Yellow
        
        # Offer alternative: save to current directory
        $saveAlternative = Read-Host "Would you like to save to the script's current directory instead? (Y/N)"
        if ($saveAlternative -eq 'y' -or $saveAlternative -eq 'Y') {
            $currentDir = Get-Location
            $exportPath = Join-Path -Path $currentDir -ChildPath $defaultFileName
            $results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
            Write-Host "Results exported to: $exportPath" -ForegroundColor Green
        }
    }
    
    # Step 10: Display example logins
    if ($successCount -gt 0) {
        Write-Host "`n--- Example User Logins ---" -ForegroundColor White
        $successfulUsers = $results | Where-Object { $_.Status -eq "Success" } | Select-Object -First 5
        foreach ($user in $successfulUsers) {
            $changeNote = if ($forceChange) { " (will be prompted to change)" } else { "" }
            Write-Host "  Username: $($user.Username) | Password: $resetPassword$changeNote" -ForegroundColor Gray
        }
        
        if ($successfulUsers.Count -lt $successCount) {
            Write-Host "  ... and $($successCount - $successfulUsers.Count) more users" -ForegroundColor Gray
        }
    }
    
    Write-Host "`nScript execution completed." -ForegroundColor Green
    Write-Host "Press Enter to exit..." -ForegroundColor Gray
    Read-Host
}

# Run the script
Main