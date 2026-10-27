# AD Inactive Account and Folder Redirection Cleanup Script
# Version 11.0 - Interactive GridView with Account Move Capability
# Description: Identifies inactive AD accounts and their redirected folders,
#              generates a CSV report, and optionally performs cleanup/move operations

param(
    [Parameter(Mandatory=$false, HelpMessage="Number of days of inactivity to check")]
    [int]$InactiveDays = 0,
    
    [Parameter(Mandatory=$false, HelpMessage="Perform account deletion after report")]
    [switch]$DeleteAccounts,
    
    [Parameter(Mandatory=$false, HelpMessage="Perform folder deletion after report")]
    [switch]$DeleteFolders,
    
    [Parameter(Mandatory=$false, HelpMessage="Move inactive accounts to another OU")]
    [switch]$MoveAccounts,
    
    [Parameter(Mandatory=$false, HelpMessage="Preview actions without making changes")]
    [switch]$WhatIf
)

# Import Active Directory module
Import-Module ActiveDirectory -ErrorAction Stop

$ErrorActionPreference = "Stop"

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Function to get last logon info
function Get-LastLogonInfo {
    param(
        [string]$SamAccountName
    )
    
    try {
        $user = Get-ADUser -Identity $SamAccountName -Properties lastLogonTimestamp, lastLogon, whenCreated -ErrorAction Stop
        
        $lltDate = if ($user.lastLogonTimestamp) {
            [DateTime]::FromFileTime($user.lastLogonTimestamp)
        } else {
            $null
        }
        
        $llDate = if ($user.lastLogon) {
            [DateTime]::FromFileTime($user.lastLogon)
        } else {
            $null
        }
        
        $mostRecent = $null
        if ($lltDate -and $llDate) {
            $mostRecent = if ($lltDate -gt $llDate) { $lltDate } else { $llDate }
        } elseif ($lltDate) {
            $mostRecent = $lltDate
        } elseif ($llDate) {
            $mostRecent = $llDate
        }
        
        return @{
            LastLogonTimestamp = $lltDate
            LastLogon = $llDate
            MostRecent = $mostRecent
            WhenCreated = $user.whenCreated
        }
    }
    catch {
        Write-ColorOutput "Warning: Could not query $SamAccountName - $_" "DarkYellow"
        return @{
            LastLogonTimestamp = $null
            LastLogon = $null
            MostRecent = $null
            WhenCreated = $null
        }
    }
}

# Function to build exclusion list
function Get-ExclusionList {
    param(
        [string]$ServiceAccountsOU
    )
    
    $exclusions = @()
    
    $builtInAccounts = @("Administrator", "Guest", "krbtgt", "DefaultAccount")
    $exclusions += $builtInAccounts
    
    if ($ServiceAccountsOU) {
        try {
            $serviceAccounts = Get-ADUser -Filter * -SearchBase $ServiceAccountsOU -Properties sAMAccountName -ErrorAction SilentlyContinue
            if ($serviceAccounts) {
                $exclusions += $serviceAccounts.sAMAccountName
            }
        }
        catch {
            Write-ColorOutput "Warning: Could not query Service Accounts OU: $_" "Yellow"
        }
    }
    
    return $exclusions | Select-Object -Unique
}

# Function to check if Folder Redirection folder exists
function Test-FolderRedirection {
    param(
        [string]$Username,
        [string]$SharePath
    )
    
    $folderPath = Join-Path -Path $SharePath -ChildPath $Username
    
    try {
        if (Test-Path -Path $folderPath) {
            $folderSize = (Get-ChildItem -Path $folderPath -Recurse -Force -ErrorAction SilentlyContinue | 
                          Measure-Object -Property Length -Sum).Sum
            return @{
                Exists = $true
                Path = $folderPath
                SizeBytes = $folderSize
                SizeMB = [Math]::Round($folderSize / 1MB, 2)
            }
        }
    }
    catch {
        # Folder doesn't exist or access denied
    }
    
    return @{
        Exists = $false
        Path = $folderPath
        SizeBytes = 0
        SizeMB = 0
    }
}

# Function to browse for a folder
function Browse-Folder {
    param(
        [string]$Description = "Select a folder",
        [string]$InitialDirectory = ""
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Description
    $folderBrowser.ShowNewFolderButton = $false
    
    if ($InitialDirectory -and (Test-Path $InitialDirectory)) {
        $folderBrowser.SelectedPath = $InitialDirectory
    }
    else {
        $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::Desktop
    }
    
    $result = $folderBrowser.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    }
    return $null
}

# Function to select an OU using Out-GridView
function Select-OU {
    param(
        [string]$Title = "Select Organizational Unit",
        [string]$Message = "Please select an OU from the list below"
    )
    
    Write-ColorOutput "`n$Message" "Cyan"
    Write-ColorOutput "Retrieving all Organizational Units from Active Directory... (this may take a moment)" "DarkYellow"
    
    try {
        # Get all OUs with useful identifying columns
        $allOUs = Get-ADOrganizationalUnit -Filter * -Properties CanonicalName, Description, Created, Modified |
            Select-Object Name, CanonicalName, Description, DistinguishedName, Created |
            Sort-Object CanonicalName
        
        if (-not $allOUs) {
            Write-ColorOutput "No OUs found in Active Directory." "Yellow"
            return $null
        }
        
        Write-ColorOutput "Found $($allOUs.Count) OUs. Use the GridView search/filter to find specific OUs." "White"
        Write-ColorOutput "Select ONE OU and click OK. Click Cancel to skip." "White"
        
        # Display OUs in a GridView for selection
        $selectedOU = $allOUs | Out-GridView -Title $Title -OutputMode Single
        
        if ($selectedOU) {
            Write-ColorOutput "Selected OU: $($selectedOU.CanonicalName)" "Green"
            return $selectedOU.DistinguishedName
        }
        else {
            Write-ColorOutput "No OU selected." "Yellow"
            return $null
        }
    }
    catch {
        Write-ColorOutput "Error retrieving OUs: $_" "Red"
        return $null
    }
}

# ============================================
# MAIN SCRIPT EXECUTION
# ============================================

Write-ColorOutput "`n=================================================================" "Cyan"
Write-ColorOutput "     AD Inactive Account and Folder Redirection Cleanup" "Cyan"
Write-ColorOutput "     Version 11.0 - Interactive Mode" "Cyan"
Write-ColorOutput "=================================================================" "Cyan"
Write-ColorOutput "Started at: $(Get-Date)" "Cyan"

# ============================================
# STEP 1: GET INACTIVE DAYS THRESHOLD
# ============================================
Write-ColorOutput "`n=== Step 1: Inactive Days Threshold ===" "Cyan"

if ($InactiveDays -eq 0) {
    do {
        $inputDays = Read-Host "Enter number of days of inactivity to check (e.g., 90, 180, 365)"
        try {
            $InactiveDays = [int]$inputDays
            if ($InactiveDays -le 0) {
                Write-ColorOutput "Please enter a positive number." "Yellow"
            }
        }
        catch {
            Write-ColorOutput "Invalid input. Please enter a number." "Yellow"
            $InactiveDays = 0
        }
    } while ($InactiveDays -le 0)
}

Write-ColorOutput "Inactive Days Threshold set to: $InactiveDays" "Green"

# ============================================
# STEP 2: GET FOLDER REDIRECTION SHARE PATH
# ============================================
Write-ColorOutput "`n=== Step 2: Folder Redirection Share ===" "Cyan"
Write-ColorOutput "You will now browse to select the root folder where user redirected folders are stored." "White"

$FolderRedirectionShare = $null
while (-not $FolderRedirectionShare) {
    $FolderRedirectionShare = Browse-Folder -Description "Select the ROOT Folder Redirection share (e.g., \\server\share\FolderRedirection)"
    
    if (-not $FolderRedirectionShare) {
        $retry = Read-Host "No folder selected. Do you want to retry? (Y/N)"
        if ($retry -ne 'Y') {
            Write-ColorOutput "Script aborted by user." "Yellow"
            Write-ColorOutput "`n" "White"
            do {
                $exitChoice = Read-Host "Type 'EXIT' to close the script"
            } while ($exitChoice -ne 'EXIT')
            return
        }
    }
}

Write-ColorOutput "Selected share path: $FolderRedirectionShare" "Green"

# Test the path
if (Test-Path -Path $FolderRedirectionShare) {
    Write-ColorOutput "Path verified: Accessible" "Green"
    try {
        $folderCount = (Get-ChildItem -Path $FolderRedirectionShare -Directory -ErrorAction Stop | Measure-Object).Count
        Write-ColorOutput "Found $folderCount existing user folders in the share" "White"
    }
    catch {
        Write-ColorOutput "Warning: Could not list folder contents" "DarkYellow"
    }
}
else {
    Write-ColorOutput "WARNING: Path does not exist or is not accessible!" "Red"
    $continue = Read-Host "Continue anyway? (Y/N)"
    if ($continue -ne 'Y') {
        Write-ColorOutput "Script aborted by user." "Yellow"
        Write-ColorOutput "`n" "White"
        do {
            $exitChoice = Read-Host "Type 'EXIT' to close the script"
        } while ($exitChoice -ne 'EXIT')
        return
    }
}

# ============================================
# STEP 3: SELECT TARGET OU TO SCAN
# ============================================
Write-ColorOutput "`n=== Step 3: Target OU Selection ===" "Cyan"
$TargetOU = $null
while (-not $TargetOU) {
    $TargetOU = Select-OU -Title "Select Target OU - Choose ONE OU to scan for inactive accounts" `
                          -Message "Select the OU containing user accounts to check for inactivity"
    
    if (-not $TargetOU) {
        $retry = Read-Host "No Target OU selected. Do you want to retry? (Y/N)"
        if ($retry -ne 'Y') {
            Write-ColorOutput "Script aborted by user." "Yellow"
            Write-ColorOutput "`n" "White"
            do {
                $exitChoice = Read-Host "Type 'EXIT' to close the script"
            } while ($exitChoice -ne 'EXIT')
            return
        }
    }
}

# ============================================
# STEP 4: OPTIONAL SERVICE ACCOUNTS OU
# ============================================
Write-ColorOutput "`n=== Step 4: Service Accounts OU (Optional) ===" "Cyan"
$ServiceAccountsOU = ""
$selectService = Read-Host "Do you want to select a Service Accounts OU to EXCLUDE from the scan? (Y/N)"

if ($selectService -eq 'Y') {
    $ServiceAccountsOU = Select-OU -Title "Select Service Accounts OU - Accounts here will be EXCLUDED" `
                                   -Message "Select the OU containing service accounts to exclude from inactivity scan"
    
    if ($ServiceAccountsOU) {
        Write-ColorOutput "Service Accounts OU set. Accounts in this OU will be excluded." "Green"
    }
    else {
        Write-ColorOutput "No Service Accounts OU selected. Only built-in accounts will be excluded." "Yellow"
    }
}
else {
    Write-ColorOutput "Skipping Service Accounts OU selection. Only built-in accounts will be excluded." "DarkYellow"
}

# ============================================
# STEP 5: CHOOSE ACTION (DELETE/MOVE/REPORT ONLY)
# ============================================
Write-ColorOutput "`n=== Step 5: Choose Action ===" "Cyan"
Write-ColorOutput "What action would you like to take with inactive accounts?" "White"
Write-ColorOutput "  1) Report Only (no changes to accounts or folders)" "White"
Write-ColorOutput "  2) Delete Inactive Accounts" "Red"
Write-ColorOutput "  3) Delete Redirected Folders" "Red"
Write-ColorOutput "  4) Delete Both Accounts AND Folders" "Red"
Write-ColorOutput "  5) Move Inactive Accounts to Another OU" "Yellow"
Write-ColorOutput "  6) Move Inactive Accounts AND Delete Folders" "Yellow"

do {
    $actionChoice = Read-Host "Enter choice (1-6)"
} while ($actionChoice -notmatch '^[1-6]$')

switch ($actionChoice) {
    '1' {
        Write-ColorOutput "Selected: Report Only" "Green"
        $DeleteAccounts = $false
        $DeleteFolders = $false
        $MoveAccounts = $false
    }
    '2' {
        Write-ColorOutput "Selected: Delete Inactive Accounts" "Red"
        $DeleteAccounts = $true
        $DeleteFolders = $false
        $MoveAccounts = $false
    }
    '3' {
        Write-ColorOutput "Selected: Delete Redirected Folders" "Red"
        $DeleteAccounts = $false
        $DeleteFolders = $true
        $MoveAccounts = $false
    }
    '4' {
        Write-ColorOutput "Selected: Delete Both Accounts AND Folders" "Red"
        $DeleteAccounts = $true
        $DeleteFolders = $true
        $MoveAccounts = $false
    }
    '5' {
        Write-ColorOutput "Selected: Move Inactive Accounts to Another OU" "Yellow"
        $DeleteAccounts = $false
        $DeleteFolders = $false
        $MoveAccounts = $true
    }
    '6' {
        Write-ColorOutput "Selected: Move Inactive Accounts AND Delete Folders" "Yellow"
        $DeleteAccounts = $false
        $DeleteFolders = $true
        $MoveAccounts = $true
    }
}

# ============================================
# STEP 5b: IF MOVING, SELECT TARGET MOVE OU
# ============================================
$TargetMoveOU = ""
if ($MoveAccounts) {
    Write-ColorOutput "`n=== Select Destination OU for Inactive Accounts ===" "Cyan"
    $TargetMoveOU = $null
    while (-not $TargetMoveOU) {
        $TargetMoveOU = Select-OU -Title "Select Destination OU - Where to MOVE inactive accounts" `
                                  -Message "Select the OU where inactive accounts will be moved to (e.g., Disabled Users OU)"
        
        if (-not $TargetMoveOU) {
            $retry = Read-Host "No destination OU selected. Do you want to retry? (Y/N)"
            if ($retry -ne 'Y') {
                Write-ColorOutput "Cannot proceed with move operation without a destination OU." "Yellow"
                Write-ColorOutput "Falling back to Report Only mode." "Yellow"
                $MoveAccounts = $false
                break
            }
        }
    }
}

# ============================================
# STEP 6: WHATIF CONFIRMATION
# ============================================
if ($DeleteAccounts -or $DeleteFolders -or $MoveAccounts) {
    Write-ColorOutput "`n=== Step 6: Safety Preview ===" "Cyan"
    $whatIfChoice = Read-Host "Do you want to run in WHATIF preview mode first? (No changes will be made) (Y/N)"
    if ($whatIfChoice -eq 'Y') {
        $WhatIf = $true
        Write-ColorOutput "WHATIF MODE ENABLED: No actual changes will be made" "Yellow"
    }
}

# ============================================
# DISPLAY CONFIGURATION SUMMARY
# ============================================
Write-ColorOutput "`n=================================================================" "Cyan"
Write-ColorOutput "     CONFIGURATION SUMMARY" "Cyan"
Write-ColorOutput "=================================================================" "Cyan"
Write-ColorOutput "Inactive Days Threshold     : $InactiveDays" "White"
Write-ColorOutput "Folder Redirection Share    : $FolderRedirectionShare" "White"
Write-ColorOutput "Target OU to Scan           : $TargetOU" "White"
if ($ServiceAccountsOU) {
    Write-ColorOutput "Service Accounts OU (Excluded): $ServiceAccountsOU" "White"
}
else {
    Write-ColorOutput "Service Accounts OU (Excluded): None" "White"
}
Write-ColorOutput "Action Mode                 : " -NoNewline
if ($DeleteAccounts -and $DeleteFolders) { Write-ColorOutput "Delete Accounts AND Folders" "Red" }
elseif ($DeleteAccounts) { Write-ColorOutput "Delete Accounts Only" "Red" }
elseif ($DeleteFolders -and $MoveAccounts) { Write-ColorOutput "Move Accounts AND Delete Folders" "Yellow" }
elseif ($MoveAccounts) { Write-ColorOutput "Move Accounts Only" "Yellow" }
else { Write-ColorOutput "Report Only" "Green" }

if ($MoveAccounts -and $TargetMoveOU) {
    Write-ColorOutput "Move Destination OU         : $TargetMoveOU" "Yellow"
}
if ($WhatIf) {
    Write-ColorOutput "WHATIF MODE                 : ENABLED (Preview Only)" "Yellow"
}
Write-ColorOutput "=================================================================" "Cyan"

# ============================================
# BUILD EXCLUSION LIST
# ============================================
$exclusionList = Get-ExclusionList -ServiceAccountsOU $ServiceAccountsOU
Write-ColorOutput "`nExcluded accounts: $($exclusionList -join ', ')" "Yellow"

# Calculate cutoff date
$cutoffDate = (Get-Date).AddDays(-$InactiveDays)
Write-ColorOutput "Cutoff date (accounts must not have logged in since): $($cutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))" "White"

# ============================================
# QUERY AND PROCESS USERS
# ============================================

Write-ColorOutput "`nQuerying Active Directory for accounts in target OU..." "White"
try {
    $allUsers = Get-ADUser -Filter {Enabled -eq $true} -SearchBase $TargetOU -Properties `
        sAMAccountName, displayName, userPrincipalName, description, department, title, `
        distinguishedName -ErrorAction Stop
}
catch {
    Write-ColorOutput "Error: Unable to query AD users. Check OU path and permissions." "Red"
    Write-ColorOutput "Error details: $_" "Red"
    Write-ColorOutput "`n" "White"
    do {
        $exitChoice = Read-Host "Type 'EXIT' to close the script"
    } while ($exitChoice -ne 'EXIT')
    return
}

Write-ColorOutput "Found $($allUsers.Count) enabled accounts in target OU" "White"

$reportData = @()
$processedCount = 0
$totalUsers = $allUsers.Count
$skippedCount = 0

foreach ($user in $allUsers) {
    $processedCount++
    
    Write-Progress -Activity "Processing user accounts" `
                   -Status "Processing $($user.sAMAccountName) ($processedCount of $totalUsers)" `
                   -PercentComplete (($processedCount / $totalUsers) * 100)
    
    # Skip excluded accounts
    if ($user.sAMAccountName -in $exclusionList) {
        Write-ColorOutput "Skipping excluded account: $($user.sAMAccountName)" "Gray"
        $skippedCount++
        continue
    }
    
    # Get logon information
    $logonInfo = Get-LastLogonInfo -SamAccountName $user.sAMAccountName
    
    $mostRecentLogon = $logonInfo.MostRecent
    $lastLogonTimestamp = $logonInfo.LastLogonTimestamp
    $lastLogon = $logonInfo.LastLogon
    
    $isInactive = $false
    $daysSinceLastLogon = $null
    $neverLoggedIn = $false
    
    if ($mostRecentLogon) {
        $daysSinceLastLogon = [Math]::Round(((Get-Date) - $mostRecentLogon).TotalDays, 2)
        if ($daysSinceLastLogon -gt $InactiveDays) {
            $isInactive = $true
        }
    }
    else {
        # No logon timestamps found - treat as inactive
        $isInactive = $true
        $neverLoggedIn = $true
        if ($logonInfo.WhenCreated) {
            $daysSinceLastLogon = [Math]::Round(((Get-Date) - $logonInfo.WhenCreated).TotalDays, 2)
        }
    }
    
    # Process inactive accounts
    if ($isInactive) {
        $folderInfo = Test-FolderRedirection -Username $user.sAMAccountName -SharePath $FolderRedirectionShare
        
        $reportEntry = [PSCustomObject]@{
            sAMAccountName = $user.sAMAccountName
            DisplayName = $user.displayName
            UserPrincipalName = $user.userPrincipalName
            DistinguishedName = $user.distinguishedName
            WhenCreated = $logonInfo.WhenCreated
            Description = $user.description
            Department = $user.department
            Title = $user.title
            LastLogonTimestamp = if ($lastLogonTimestamp) { $lastLogonTimestamp.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
            LastLogonCurrentDC = if ($lastLogon) { $lastLogon.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
            MostRecentLogon = if ($mostRecentLogon) { $mostRecentLogon.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
            DaysSinceLastLogon = $daysSinceLastLogon
            NeverLoggedIn = $neverLoggedIn
            IsInactive = $true
            FolderRedirectionExists = $folderInfo.Exists
            FolderRedirectionPath = $folderInfo.Path
            FolderSizeMB = $folderInfo.SizeMB
            WillBeDeleted = if ($DeleteAccounts) { "Yes" } else { "No" }
            WillBeMoved = if ($MoveAccounts) { "Yes" } else { "No" }
            MoveDestination = if ($MoveAccounts) { $TargetMoveOU } else { "" }
            FolderWillBeDeleted = if ($DeleteFolders -and $folderInfo.Exists) { "Yes" } else { "No" }
        }
        
        $reportData += $reportEntry
        
        $statusColor = if ($folderInfo.Exists) { "Magenta" } else { "Yellow" }
        $folderStatus = if ($folderInfo.Exists) { "Found ($($folderInfo.SizeMB) MB)" } else { "Not Found" }
        if ($neverLoggedIn) {
            Write-ColorOutput "INACTIVE (Never Logged In): $($user.sAMAccountName) - Folder: $folderStatus" $statusColor
        }
        else {
            $logonDisplay = $mostRecentLogon.ToString("yyyy-MM-dd")
            Write-ColorOutput "INACTIVE: $($user.sAMAccountName) - Last logon: $logonDisplay - Folder: $folderStatus" $statusColor
        }
    }
}

Write-Progress -Activity "Processing user accounts" -Completed

# ============================================
# PROCESSING SUMMARY
# ============================================

$inactiveCount = $reportData.Count
$neverLoggedInCount = ($reportData | Where-Object { $_.NeverLoggedIn -eq $true }).Count
$exceededDaysCount = $inactiveCount - $neverLoggedInCount
$activeCount = $totalUsers - $skippedCount - $inactiveCount

Write-ColorOutput "`n=================================================================" "Cyan"
Write-ColorOutput "     PROCESSING SUMMARY" "Cyan"
Write-ColorOutput "=================================================================" "Cyan"
Write-ColorOutput "Total users in OU           : $totalUsers" "White"
Write-ColorOutput "Skipped (excluded)          : $skippedCount" "Gray"
Write-ColorOutput "Users processed             : $($totalUsers - $skippedCount)" "White"
Write-ColorOutput "Active accounts             : $activeCount" "Green"
Write-ColorOutput "Inactive accounts found     : $inactiveCount" "Yellow"
Write-ColorOutput "  - Never logged in         : $neverLoggedInCount" "Yellow"
Write-ColorOutput "  - Exceeded $InactiveDays days    : $exceededDaysCount" "Yellow"

# ============================================
# FOLDER SUMMARY AND REPORT
# ============================================

if ($reportData.Count -gt 0) {
    $accountsWithFolders = ($reportData | Where-Object { $_.FolderRedirectionExists -eq $true }).Count
    $accountsWithoutFolders = ($reportData | Where-Object { $_.FolderRedirectionExists -eq $false }).Count
    $totalFolderSizeMB = ($reportData | Where-Object { $_.FolderRedirectionExists -eq $true } | Measure-Object -Property FolderSizeMB -Sum).Sum
    
    Write-ColorOutput "`n=================================================================" "Cyan"
    Write-ColorOutput "     FOLDER SUMMARY" "Cyan"
    Write-ColorOutput "=================================================================" "Cyan"
    Write-ColorOutput "Inactive accounts WITH redirected folders    : $accountsWithFolders" "White"
    Write-ColorOutput "Inactive accounts WITHOUT redirected folders : $accountsWithoutFolders" "White"
    Write-ColorOutput "Total folder space used                      : $([Math]::Round($totalFolderSizeMB, 2)) MB" "White"
    
    Write-ColorOutput "`n=================================================================" "Cyan"
    Write-ColorOutput "     INACTIVE ACCOUNTS FOUND" "Cyan"
    Write-ColorOutput "=================================================================" "Cyan"
    $reportData | Format-Table sAMAccountName, DisplayName, NeverLoggedIn, DaysSinceLastLogon, FolderRedirectionExists, FolderSizeMB -AutoSize
    
    # Ask user if they want to save the report
    Write-ColorOutput "`n=== Save Report ===" "Cyan"
    $saveReport = Read-Host "Do you want to save the report to a CSV file? (Y/N)"
    
    if ($saveReport -eq 'Y') {
        $defaultFileName = "InactiveAccounts_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $defaultPath = Join-Path -Path [Environment]::GetFolderPath('Desktop') -ChildPath $defaultFileName
        
        Write-ColorOutput "Default save location: $defaultPath" "Gray"
        $savePath = Read-Host "Enter save path (press Enter to use default)"
        if (-not $savePath) {
            $savePath = $defaultPath
        }
        
        try {
            $reportData | Export-Csv -Path $savePath -NoTypeInformation
            Write-ColorOutput "Report saved to: $savePath" "Green"
        }
        catch {
            Write-ColorOutput "Error saving report: $_" "Red"
        }
    }
    else {
        Write-ColorOutput "Report not saved." "Yellow"
    }
    
    # ============================================
    # ACTION: MOVE OR DELETE OPERATIONS
    # ============================================
    
    if ($MoveAccounts -or $DeleteAccounts -or $DeleteFolders) {
        Write-ColorOutput "`n=================================================================" "Red"
        Write-ColorOutput "     WARNING: CHANGES WILL BE MADE" "Red"
        Write-ColorOutput "=================================================================" "Red"
        
        if ($WhatIf) {
            Write-ColorOutput "WHATIF MODE: No actual changes will be made - PREVIEW ONLY" "Yellow"
        }
        
        $confirmMessage = "`nYou are about to perform the following actions:`n"
        if ($MoveAccounts) {
            $confirmMessage += "  - MOVE $($reportData.Count) inactive accounts to: $TargetMoveOU`n"
        }
        if ($DeleteAccounts) {
            $confirmMessage += "  - DELETE $($reportData.Count) inactive user accounts from Active Directory`n"
        }
        if ($DeleteFolders) {
            $foldersToDelete = ($reportData | Where-Object { $_.FolderRedirectionExists -eq $true }).Count
            $confirmMessage += "  - DELETE $foldersToDelete redirected folders`n"
        }
        $confirmMessage += "`nAre you sure you want to proceed? (Type 'YES' to confirm)"
        
        Write-ColorOutput $confirmMessage "Red"
        $confirmation = Read-Host
        
        if ($confirmation -eq "YES") {
            Write-ColorOutput "`nProceeding with operations..." "Red"
            
            $successCount = 0
            $failureCount = 0
            
            foreach ($account in $reportData) {
                # Move accounts if requested
                if ($MoveAccounts) {
                    try {
                        if (-not $WhatIf) {
                            # Disable the account first, then move it
                            Disable-ADAccount -Identity $account.sAMAccountName -Confirm:$false -ErrorAction Stop
                            Move-ADObject -Identity $account.DistinguishedName -TargetPath $TargetMoveOU -ErrorAction Stop
                            # Update the account description with move details
                            $newDescription = "Moved from: $($account.DistinguishedName) on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Original: $($account.Description)"
                            Set-ADUser -Identity $account.sAMAccountName -Description $newDescription -ErrorAction SilentlyContinue
                        }
                        Write-ColorOutput "Account moved and disabled: $($account.sAMAccountName) -> $TargetMoveOU" "Yellow"
                        $successCount++
                    }
                    catch {
                        Write-ColorOutput "Failed to move account: $($account.sAMAccountName) - $_" "DarkRed"
                        $failureCount++
                    }
                }
                
                # Delete accounts if requested
                if ($DeleteAccounts) {
                    try {
                        if (-not $WhatIf) {
                            Remove-ADUser -Identity $account.sAMAccountName -Confirm:$false -ErrorAction Stop
                        }
                        Write-ColorOutput "Account deleted: $($account.sAMAccountName)" "Red"
                        $successCount++
                    }
                    catch {
                        Write-ColorOutput "Failed to delete account: $($account.sAMAccountName) - $_" "DarkRed"
                        $failureCount++
                    }
                }
                
                # Delete folders if requested
                if ($DeleteFolders -and $account.FolderRedirectionExists) {
                    try {
                        if (-not $WhatIf) {
                            Remove-Item -Path $account.FolderRedirectionPath -Recurse -Force -ErrorAction Stop
                        }
                        Write-ColorOutput "Folder deleted: $($account.FolderRedirectionPath)" "Red"
                        $successCount++
                    }
                    catch {
                        Write-ColorOutput "Failed to delete folder: $($account.FolderRedirectionPath) - $_" "DarkRed"
                        $failureCount++
                    }
                }
            }
            
            Write-ColorOutput "`n=================================================================" "Green"
            Write-ColorOutput "     OPERATIONS COMPLETE" "Green"
            Write-ColorOutput "=================================================================" "Green"
            Write-ColorOutput "Successful operations: $successCount" "Green"
            if ($failureCount -gt 0) {
                Write-ColorOutput "Failed operations    : $failureCount" "Red"
            }
        }
        else {
            Write-ColorOutput "`nOperations cancelled. No changes were made." "Green"
        }
    }
}
else {
    Write-ColorOutput "`nNo inactive accounts found in the specified OU." "Green"
}

Write-ColorOutput "`n=================================================================" "Cyan"
Write-ColorOutput "     Script Completed at: $(Get-Date)" "Cyan"
Write-ColorOutput "=================================================================" "Cyan"

# ============================================
# MANUAL EXIT TRIGGER
# ============================================
Write-ColorOutput "`n" "White"
Write-ColorOutput "Script execution finished. Review the output above." "Cyan"
do {
    $exitChoice = Read-Host "Type 'EXIT' to close the script"
} while ($exitChoice -ne 'EXIT')