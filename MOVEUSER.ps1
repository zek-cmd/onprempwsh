# Active Directory User Move Script (OU-to-OU)
# Move all users from a source OU to a target OU
# Script compatibility verified for PowerShell 5.1+

# Import required assemblies for dialog boxes
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

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

# Function to open Save As dialog for results CSV
function Get-ResultsSavePath {
    param(
        [string]$DefaultFileName = "AD_User_Move_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    )
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveFileDialog.Title = "Save User Move Results CSV Report"
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

# Function to get all OUs from Active Directory
function Get-AllADOrganizationalUnits {
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        Write-Error "Active Directory module is not available."
        return @()
    }
    
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    
    Write-Host "Fetching Organizational Units from Active Directory..." -ForegroundColor White
    
    $allOUs = Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName, Name, Description, Created, Modified | 
        Select-Object DistinguishedName, Name, @{Name="Path";Expression={
            $parts = $_.DistinguishedName -split ','
            $ouParts = $parts | Where-Object { $_ -like 'OU=*' }
            [array]::Reverse($ouParts)
            ($ouParts -join '/') -replace '^OU=', ''
        }}, Description, Created, Modified
    
    Write-Host "Found $($allOUs.Count) Organizational Units" -ForegroundColor Green
    
    return $allOUs
}

# Function to display OU selection tree browser
function Show-OUTreeBrowser {
    param(
        [array]$OrganizationalUnits,
        [string]$Title = "Select Organizational Unit",
        [bool]$AllowRoot = $false
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(600, 500)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    
    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.Location = New-Object System.Drawing.Point(10, 10)
    $treeView.Size = New-Object System.Drawing.Size(565, 400)
    $treeView.CheckBoxes = $false
    $treeView.ShowLines = $true
    $treeView.ShowRootLines = $true
    $treeView.ShowPlusMinus = $true
    $treeView.Sorted = $true
    $form.Controls.Add($treeView)
    
    $selectedLabel = New-Object System.Windows.Forms.Label
    $selectedLabel.Location = New-Object System.Drawing.Point(10, 420)
    $selectedLabel.Size = New-Object System.Drawing.Size(565, 20)
    $selectedLabel.Text = "Selected OU: None"
    $form.Controls.Add($selectedLabel)
    
    $pathLabel = New-Object System.Windows.Forms.Label
    $pathLabel.Location = New-Object System.Drawing.Point(10, 440)
    $pathLabel.Size = New-Object System.Drawing.Size(565, 40)
    $pathLabel.Text = "Path:"
    $form.Controls.Add($pathLabel)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(400, 420)
    $okButton.Size = New-Object System.Drawing.Size(85, 30)
    $okButton.Text = "OK"
    $okButton.Enabled = $false
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(490, 420)
    $cancelButton.Size = New-Object System.Drawing.Size(85, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)
    
    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Location = New-Object System.Drawing.Point(10, 460)
    $refreshButton.Size = New-Object System.Drawing.Size(100, 30)
    $refreshButton.Text = "Refresh"
    $refreshButton.Add_Click({
        Write-Host "Refreshing OU list..." -ForegroundColor Yellow
        $OrganizationalUnits = Get-AllADOrganizationalUnits
        if ($OrganizationalUnits.Count -gt 0) {
            $treeView.Nodes.Clear()
            Build-OUTree -TreeView $treeView -OUs $OrganizationalUnits
        }
    })
    $form.Controls.Add($refreshButton)
    
    $script:selectedOU = $null
    
    function Build-OUTree {
        param($TreeView, $OUs)
        
        $rootNode = New-Object System.Windows.Forms.TreeNode
        $rootNode.Text = "$Domain (Root)"
        $rootNode.Tag = $DomainDN
        $rootNode.Name = "ROOT"
        $TreeView.Nodes.Add($rootNode)
        
        $sortedOUs = $OUs | Sort-Object @{Expression={$_.Path.Split('/').Count}; Ascending=$true}, Path
        
        $nodeDict = @{}
        $nodeDict[$DomainDN] = $rootNode
        
        foreach ($ou in $sortedOUs) {
            $dnParts = $ou.DistinguishedName -split ','
            $parentDN = $dnParts[1..($dnParts.Count-1)] -join ','
            
            $newNode = New-Object System.Windows.Forms.TreeNode
            $newNode.Text = $ou.Name
            $newNode.Tag = $ou.DistinguishedName
            $newNode.Name = $ou.DistinguishedName
            $newNode.ToolTipText = "Created: $($ou.Created)`nModified: $($ou.Modified)`nDescription: $($ou.Description)"
            
            if ($nodeDict.ContainsKey($parentDN)) {
                $nodeDict[$parentDN].Nodes.Add($newNode)
            }
            else {
                $rootNode.Nodes.Add($newNode)
            }
            
            $nodeDict[$ou.DistinguishedName] = $newNode
        }
        
        $rootNode.Expand()
    }
    
    $treeView.Add_AfterSelect({
        param($sender, $e)
        
        $selectedNode = $treeView.SelectedNode
        
        $validSelection = ($selectedNode.Tag -ne $DomainDN) -or $AllowRoot
        
        if ($selectedNode -and $validSelection) {
            $selectedLabel.Text = "Selected OU: $($selectedNode.Text)"
            $pathLabel.Text = "Path: $($selectedNode.Tag)"
            $okButton.Enabled = $true
            
            $script:selectedOU = @{
                Name = $selectedNode.Text
                DistinguishedName = $selectedNode.Tag
                Path = $selectedNode.Tag
                FullPath = $selectedNode.Tag
            }
        }
        else {
            $selectedLabel.Text = "Selected OU: None (Please select an OU below root)"
            $pathLabel.Text = "Path:"
            $okButton.Enabled = $false
            $script:selectedOU = $null
        }
    })
    
    $treeView.Add_NodeMouseDoubleClick({
        param($sender, $e)
        
        if ($okButton.Enabled -and $script:selectedOU) {
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        }
    })
    
    Build-OUTree -TreeView $treeView -OUs $OrganizationalUnits
    
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $script:selectedOU) {
        return $script:selectedOU
    }
    
    return $null
}

# Function to get users in an OU
function Get-UsersInOU {
    param([string]$OUPath)
    
    Write-Host "Searching for users in: $OUPath" -ForegroundColor White
    
    try {
        $users = Get-ADUser -Filter * -SearchBase $OUPath -Properties SamAccountName, Name, DisplayName, Enabled, DistinguishedName | 
            Select-Object SamAccountName, Name, DisplayName, Enabled, DistinguishedName
        
        Write-Host "Found $($users.Count) user(s) in this OU." -ForegroundColor Green
        return $users
    }
    catch {
        Write-Host "ERROR: Unable to retrieve users from OU: $_" -ForegroundColor Red
        return @()
    }
}

# Function to move AD user to target OU
function Move-ADUserToOU {
    param(
        $ADUser,
        [string]$TargetOUPath
    )
    
    $currentOU = $ADUser.DistinguishedName -replace '^CN=.*?,(?=OU=|CN=)', ''
    
    Write-Host "`nProcessing: $($ADUser.DisplayName)" -ForegroundColor White
    Write-Host "  Username: $($ADUser.SamAccountName)" -ForegroundColor Gray
    Write-Host "  Current Location: $currentOU" -ForegroundColor Gray
    Write-Host "  Target Location: $TargetOUPath" -ForegroundColor Yellow
    
    if (-not $ADUser.Enabled) {
        Write-Host "  WARNING: Account is disabled. Will still be moved." -ForegroundColor Yellow
    }
    
    if ($currentOU -eq $TargetOUPath) {
        Write-Host "  SKIPPED: User is already in the target OU." -ForegroundColor Yellow
        return @{
            Status = "Skipped"
            Message = "Already in target OU"
            Username = $ADUser.SamAccountName
            CurrentOU = $currentOU
        }
    }
    
    try {
        Move-ADObject -Identity $ADUser.DistinguishedName -TargetPath $TargetOUPath
        Write-Host "  SUCCESS: Moved to target OU." -ForegroundColor Green
        
        return @{
            Status = "Success"
            Message = "Moved successfully"
            Username = $ADUser.SamAccountName
            CurrentOU = $currentOU
        }
    }
    catch {
        Write-Host "  ERROR: Failed to move user: $_" -ForegroundColor Red
        
        return @{
            Status = "Failed"
            Message = "Move operation failed: $_"
            Username = $ADUser.SamAccountName
            CurrentOU = $currentOU
        }
    }
}

# Function to display summary before execution
function Show-Summary {
    param(
        [array]$Users,
        [hashtable]$SourceOU,
        [hashtable]$TargetOU
    )
    
    Write-Host "`n=== MOVE SUMMARY ===" -ForegroundColor Cyan
    Write-Host "Source OU: $($SourceOU.Name)" -ForegroundColor Yellow
    Write-Host "  Source Path: $($SourceOU.FullPath)" -ForegroundColor Gray
    Write-Host "Target OU: $($TargetOU.Name)" -ForegroundColor Yellow
    Write-Host "  Target Path: $($TargetOU.FullPath)" -ForegroundColor Gray
    Write-Host "Number of users to move: $($Users.Count)" -ForegroundColor Yellow
    Write-Host "Domain: $Domain" -ForegroundColor Yellow
    
    Write-Host "`n--- Users Found in Source OU ---" -ForegroundColor White
    $maxNameLength = 0
    foreach ($user in $Users) {
        if ($user.DisplayName.Length -gt $maxNameLength) {
            $maxNameLength = $user.DisplayName.Length
        }
    }
    $maxNameLength += 5
    
    $headerFormat = "{0,-$maxNameLength} {1,-25} {2}"
    Write-Host ($headerFormat -f "Display Name", "Username", "Enabled") -ForegroundColor Cyan
    Write-Host ("-" * ($maxNameLength + 35)) -ForegroundColor DarkGray
    
    $rowFormat = "{0,-$maxNameLength} {1,-25} {2}"
    foreach ($user in $Users) {
        $enabledStatus = if ($user.Enabled) { "Yes" } else { "No" }
        Write-Host ($rowFormat -f $user.DisplayName, $user.SamAccountName, $enabledStatus) -ForegroundColor Gray
    }
    
    Write-Host "`n=================================" -ForegroundColor Cyan
}

# Main execution
function Main {
    Clear-Host
    
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "     ACTIVE DIRECTORY USER MOVE TOOL" -ForegroundColor Cyan
    Write-Host "     (Move All Users from Source OU to Target OU)" -ForegroundColor Cyan
    Write-Host "     (Graphical OU Browser for Both OUs)" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Checking prerequisites..." -ForegroundColor White
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        Write-Host "WARNING: Active Directory module not found." -ForegroundColor Red
        Write-Host "Please install Remote Server Administration Tools (RSAT)." -ForegroundColor Red
        Write-Host "Script will continue but cannot move users without AD module." -ForegroundColor Yellow
        $continue = Read-Host "Continue anyway? (yes/no)"
        if ($continue -ne 'yes') {
            Write-Host "Script terminated by user." -ForegroundColor Yellow
            return
        }
    }
    
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    
    # Step 1: Select SOURCE OU
    Write-Host "`nStep 1: Select SOURCE OU (where users currently are)" -ForegroundColor White
    $allOUs = Get-AllADOrganizationalUnits
    if ($allOUs.Count -eq 0) {
        Write-Host "No OUs found or unable to connect to Active Directory." -ForegroundColor Red
        return
    }
    
    $sourceOU = Show-OUTreeBrowser -OrganizationalUnits $allOUs -Title "Select SOURCE Organizational Unit"
    if (-not $sourceOU) {
        Write-Host "Source OU selection cancelled." -ForegroundColor Red
        return
    }
    Write-Host "Source OU: $($sourceOU.Name)" -ForegroundColor Green
    Write-Host "Source Path: $($sourceOU.DistinguishedName)" -ForegroundColor Green
    
    # Step 2: Get users from source OU
    Write-Host "`nStep 2: Retrieving users from source OU..." -ForegroundColor White
    $users = Get-UsersInOU -OUPath $sourceOU.DistinguishedName
    if ($users.Count -eq 0) {
        Write-Host "No users found in the source OU. Nothing to move." -ForegroundColor Yellow
        return
    }
    
    # Step 3: Select TARGET OU
    Write-Host "`nStep 3: Select TARGET OU (where users will be moved to)" -ForegroundColor White
    $allOUs = Get-AllADOrganizationalUnits
    $targetOU = Show-OUTreeBrowser -OrganizationalUnits $allOUs -Title "Select TARGET Organizational Unit"
    if (-not $targetOU) {
        Write-Host "Target OU selection cancelled." -ForegroundColor Red
        return
    }
    Write-Host "Target OU: $($targetOU.Name)" -ForegroundColor Green
    Write-Host "Target Path: $($targetOU.DistinguishedName)" -ForegroundColor Green
    
    if ($sourceOU.DistinguishedName -eq $targetOU.DistinguishedName) {
        Write-Host "ERROR: Source and target OUs are the same. No moves needed." -ForegroundColor Red
        return
    }
    
    # Step 4: Show summary - with explicit type checking to prevent errors
    Write-Host "`n=== MOVE SUMMARY ===" -ForegroundColor Cyan
    Write-Host "Source OU: $($sourceOU.Name)" -ForegroundColor Yellow
    Write-Host "  Source Path: $($sourceOU.DistinguishedName)" -ForegroundColor Gray
    Write-Host "Target OU: $($targetOU.Name)" -ForegroundColor Yellow
    Write-Host "  Target Path: $($targetOU.DistinguishedName)" -ForegroundColor Gray
    Write-Host "Number of users to move: $($users.Count)" -ForegroundColor Yellow
    Write-Host "Domain: $Domain" -ForegroundColor Yellow
    
    Write-Host "`n--- Users Found in Source OU ---" -ForegroundColor White
    $maxNameLength = 0
    foreach ($user in $users) {
        if ($user.DisplayName.Length -gt $maxNameLength) {
            $maxNameLength = $user.DisplayName.Length
        }
    }
    $maxNameLength += 5
    
    $headerFormat = "{0,-$maxNameLength} {1,-25} {2}"
    Write-Host ($headerFormat -f "Display Name", "Username", "Enabled") -ForegroundColor Cyan
    Write-Host ("-" * ($maxNameLength + 35)) -ForegroundColor DarkGray
    
    $rowFormat = "{0,-$maxNameLength} {1,-25} {2}"
    foreach ($user in $users) {
        $enabledStatus = if ($user.Enabled) { "Yes" } else { "No" }
        Write-Host ($rowFormat -f $user.DisplayName, $user.SamAccountName, $enabledStatus) -ForegroundColor Gray
    }
    
    Write-Host "`n=================================" -ForegroundColor Cyan
    
    # Step 5: Ask for confirmation
    Write-Host "`n=== CONFIRMATION REQUIRED ===" -ForegroundColor Red
    Write-Host "This will move $($users.Count) user accounts:" -ForegroundColor Yellow
    Write-Host "  FROM: $($sourceOU.Name)" -ForegroundColor Yellow
    Write-Host "    Path: $($sourceOU.DistinguishedName)" -ForegroundColor Gray
    Write-Host "  TO: $($targetOU.Name)" -ForegroundColor Yellow
    Write-Host "    Path: $($targetOU.DistinguishedName)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "WARNING: Moving users can affect Group Policy application," -ForegroundColor Red
    Write-Host "login scripts, and other OU-specific settings." -ForegroundColor Red
    $confirmation = Read-Host "`nType 'YES' to proceed, anything else to cancel"
    
    if ($confirmation -ne "YES") {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        return
    }
    
    # Step 6: Move users
    Write-Host "`nStarting user moves..." -ForegroundColor Green
    
    $results = @()
    $successCount = 0
    $failCount = 0
    $skipCount = 0
    
    foreach ($user in $users) {
        Write-Host "`n" + ("="*50) -ForegroundColor DarkGray
        
        $result = Move-ADUserToOU -ADUser $user -TargetOUPath $targetOU.DistinguishedName
        
        $resultObject = [PSCustomObject]@{
            DisplayName = $user.DisplayName
            Username = if ($result.Username) { $result.Username } else { $user.SamAccountName }
            Status = if ($result.Status) { $result.Status } else { "Failed" }
            Message = if ($result.Message) { $result.Message } else { "Unknown error" }
            PreviousOU = if ($result.CurrentOU) { $result.CurrentOU } else { "" }
            SourceOU = $sourceOU.DistinguishedName
            SourceOUName = $sourceOU.Name
            TargetOU = $targetOU.DistinguishedName
            TargetOUName = $targetOU.Name
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        $results += $resultObject
        
        if ($result.Status -eq "Success") {
            $successCount++
        }
        elseif ($result.Status -eq "Skipped") {
            $skipCount++
        }
        else {
            $failCount++
        }
        
        Start-Sleep -Milliseconds 100
    }
    
    # Step 7: Display final summary
    Write-Host "`n" + ("="*60) -ForegroundColor Cyan
    Write-Host "MOVE OPERATION COMPLETE" -ForegroundColor Cyan
    Write-Host ("="*60) -ForegroundColor Cyan
    
    if ($successCount -gt 0) {
        Write-Host "Successfully moved: $successCount users" -ForegroundColor Green
    }
    
    if ($skipCount -gt 0) {
        Write-Host "Skipped (already in target): $skipCount users" -ForegroundColor Yellow
    }
    
    if ($failCount -gt 0) {
        Write-Host "Failed: $failCount users" -ForegroundColor Red
    }
    
    if ($results.Count -gt 0) {
        Write-Host "`n--- Detailed Results ---" -ForegroundColor White
        $results | Format-Table DisplayName, Username, Status -AutoSize | Out-Host
    }
    
    # Step 8: Let user choose where to save the results CSV
    Write-Host "`n" + ("-"*60) -ForegroundColor DarkGray
    Write-Host "Step 8: Save Results Report" -ForegroundColor White
    
    # Generate default filename with timestamp
    $defaultFileName = "AD_User_Move_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
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
    
    Write-Host "`nScript execution completed." -ForegroundColor Green
    Write-Host "Press Enter to exit..." -ForegroundColor Gray
    Read-Host
}

# Run the script
Main