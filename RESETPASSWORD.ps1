# Active Directory Password Reset Script (OU-Based)
# Reset passwords for all users in a selected Organizational Unit
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
    param([array]$OrganizationalUnits)
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Target Organizational Unit"
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
        if ($selectedNode -and $selectedNode.Tag -ne $DomainDN) {
            $selectedLabel.Text = "Selected OU: $($selectedNode.Text)"
            $pathLabel.Text = "Path: $($selectedNode.Tag)"
            $okButton.Enabled = $true
            
            $script:selectedOU = @{
                Name = $selectedNode.Text
                DistinguishedName = $selectedNode.Tag
                Path = $selectedNode.Tag
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

# Function to get target OU using tree browser
function Get-TargetOU {
    Write-Host "`n=== Target OU Selection ===" -ForegroundColor Cyan
    Write-Host "Current domain: $Domain" -ForegroundColor Yellow
    Write-Host "Loading OU tree browser..." -ForegroundColor Yellow
    
    $allOUs = Get-AllADOrganizationalUnits
    
    if ($allOUs.Count -eq 0) {
        Write-Host "No OUs found or unable to connect to Active Directory." -ForegroundColor Red
        Write-Host "Falling back to manual OU entry..." -ForegroundColor Yellow
        
        return Get-TargetOUMannualy
    }
    
    $selectedOU = Show-OUTreeBrowser -OrganizationalUnits $allOUs
    
    if (-not $selectedOU) {
        Write-Host "OU selection cancelled by user." -ForegroundColor Yellow
        
        $manualChoice = Read-Host "Do you want to enter the OU path manually? (Y/N)"
        if ($manualChoice -eq 'Y' -or $manualChoice -eq 'y') {
            return Get-TargetOUMannualy
        }
        return $null
    }
    
    Write-Host "Selected Target OU: $($selectedOU.Name)" -ForegroundColor Green
    Write-Host "OU Path: $($selectedOU.DistinguishedName)" -ForegroundColor Green
    
    $ouParts = $selectedOU.DistinguishedName -split ','
    $ouComponents = $ouParts | Where-Object { $_ -like 'OU=*' }
    
    $rootOU = ""
    $subOU = ""
    
    if ($ouComponents.Count -ge 1) {
        $rootOU = ($ouComponents[0] -replace 'OU=', '')
    }
    if ($ouComponents.Count -ge 2) {
        $subOU = ($ouComponents[1] -replace 'OU=', '')
    }
    
    return @{
        RootOU = $rootOU
        SubOU = $subOU
        FullPath = $selectedOU.DistinguishedName
        Name = $selectedOU.Name
    }
}

# Fallback function for manual OU entry
function Get-TargetOUMannualy {
    Write-Host "`n=== Manual OU Entry ===" -ForegroundColor Cyan
    Write-Host "Current domain: $Domain" -ForegroundColor Yellow
    Write-Host "Domain DN: $DomainDN" -ForegroundColor Yellow
    Write-Host "Default path format: OU=SubOU,OU=RootOU,$DomainDN" -ForegroundColor Yellow
    
    $rootOU = Read-Host "Enter Root OU name (e.g., SITC)"
    if ([string]::IsNullOrWhiteSpace($rootOU)) {
        $rootOU = "SITC"
    }
    
    $subOU = Read-Host "Enter Sub OU name (e.g., TESTING)"
    if ([string]::IsNullOrWhiteSpace($subOU)) {
        $subOU = "TESTING"
    }
    
    $rootOU = $rootOU -replace '[\/\\#]', ''
    $subOU = $subOU -replace '[\/\\#]', ''
    
    $ouPath = "OU=$subOU,OU=$rootOU,$DomainDN"
    
    Write-Host "Selected OU Path: $ouPath" -ForegroundColor Green
    
    return @{
        RootOU = $rootOU
        SubOU = $subOU
        FullPath = $ouPath
        Name = "$subOU ($rootOU)"
    }
}

# Function to get all users in an OU
function Get-UsersInOU {
    param([string]$OUPath)
    
    Write-Host "Searching for users in: $OUPath" -ForegroundColor White
    
    try {
        $users = Get-ADUser -Filter * -SearchBase $OUPath -Properties SamAccountName, Name, DisplayName, Enabled, PasswordLastSet | 
            Select-Object SamAccountName, Name, DisplayName, Enabled, PasswordLastSet, DistinguishedName
        
        Write-Host "Found $($users.Count) user(s) in this OU." -ForegroundColor Green
        return $users
    }
    catch {
        Write-Host "ERROR: Unable to retrieve users from OU: $_" -ForegroundColor Red
        return @()
    }
}

# Function to reset password for an AD user
function Reset-ADUserPassword {
    param(
        $ADUser,
        [string]$Password
    )
    
    Write-Host "`nProcessing: $($ADUser.DisplayName)" -ForegroundColor White
    Write-Host "  Username: $($ADUser.SamAccountName)" -ForegroundColor Gray
    Write-Host "  Account Enabled: $($ADUser.Enabled)" -ForegroundColor Gray
    
    if ($ADUser.PasswordLastSet) {
        Write-Host "  Password Last Set: $($ADUser.PasswordLastSet)" -ForegroundColor Gray
    }
    else {
        Write-Host "  Password Last Set: Never" -ForegroundColor Gray
    }
    
    if (-not $ADUser.Enabled) {
        Write-Host "  WARNING: Account is disabled. Password will still be reset." -ForegroundColor Yellow
    }
    
    try {
        $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        Set-ADAccountPassword -Identity $ADUser.DistinguishedName -NewPassword $securePassword -Reset
        
        Write-Host "  SUCCESS: Password reset to '$Password'." -ForegroundColor Green
        
        return @{
            Status = "Success"
            Message = "Password reset successfully"
            Username = $ADUser.SamAccountName
            DisplayName = $ADUser.DisplayName
        }
    }
    catch {
        Write-Host "  ERROR: Failed to reset password: $_" -ForegroundColor Red
        
        return @{
            Status = "Failed"
            Message = "Password reset failed: $_"
            Username = $ADUser.SamAccountName
            DisplayName = $ADUser.DisplayName
        }
    }
}

# Function to display summary before execution
function Show-Summary {
    param(
        [array]$Users,
        [hashtable]$OUInfo,
        [string]$Password
    )
    
    Write-Host "`n=== PASSWORD RESET SUMMARY ===" -ForegroundColor Cyan
    Write-Host "Target OU: $($OUInfo.Name)" -ForegroundColor Yellow
    Write-Host "  OU Path: $($OUInfo.FullPath)" -ForegroundColor Yellow
    Write-Host "Number of users to process: $($Users.Count)" -ForegroundColor Yellow
    Write-Host "New Password: $Password" -ForegroundColor Yellow
    Write-Host "Domain: $Domain" -ForegroundColor Yellow
    
    Write-Host "`n--- Users Found in OU ---" -ForegroundColor White
    $maxNameLength = 30
    $formatString = "{0,-$maxNameLength} {1,-25} {2}"
    Write-Host ($formatString -f "Display Name", "Username", "Enabled") -ForegroundColor Cyan
    Write-Host ("-" * ($maxNameLength + 35)) -ForegroundColor DarkGray
    
    foreach ($user in $Users) {
        $enabledStatus = if ($user.Enabled) { "Yes" } else { "No" }
        Write-Host ($formatString -f $user.DisplayName, $user.SamAccountName, $enabledStatus) -ForegroundColor Gray
    }
    
    Write-Host "`n=================================" -ForegroundColor Cyan
}

# Main execution
function Main {
    Clear-Host
    
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "     ACTIVE DIRECTORY PASSWORD RESET TOOL" -ForegroundColor Cyan
    Write-Host "     (OU-Based Bulk Password Reset)" -ForegroundColor Cyan
    Write-Host "     (Graphical OU Browser)" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Check for required modules
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
    
    # Step 1: Get target OU using tree browser
    Write-Host "`nStep 1: Select target OU containing users" -ForegroundColor White
    $ouInfo = Get-TargetOU
    if (-not $ouInfo) {
        Write-Host "OU selection failed or cancelled." -ForegroundColor Red
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($ouInfo.FullPath)) {
        Write-Host "ERROR: OU path is empty. Cannot proceed." -ForegroundColor Red
        return
    }
    
    # Step 2: Get all users in the selected OU
    Write-Host "`nStep 2: Retrieving users from selected OU..." -ForegroundColor White
    $users = Get-UsersInOU -OUPath $ouInfo.FullPath
    
    if ($users.Count -eq 0) {
        Write-Host "No users found in the selected OU." -ForegroundColor Red
        Write-Host "Make sure the OU contains user accounts." -ForegroundColor Yellow
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
    Show-Summary -Users $users -OUInfo $ouInfo -Password $resetPassword
    
    # Step 5: Ask for confirmation
    Write-Host "`n=== CONFIRMATION REQUIRED ===" -ForegroundColor Red
    Write-Host "This will reset passwords for $($users.Count) users to: '$resetPassword'" -ForegroundColor Yellow
    Write-Host "  OU: $($ouInfo.Name)" -ForegroundColor Yellow
    Write-Host "  OU Path: $($ouInfo.FullPath)" -ForegroundColor Gray
    Write-Host "  Domain: $Domain" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "WARNING: This will forcefully change ALL user passwords in this OU." -ForegroundColor Red
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
    
    foreach ($user in $users) {
        Write-Host "`n" + ("="*50) -ForegroundColor DarkGray
        
        $result = Reset-ADUserPassword -ADUser $user -Password $resetPassword
        
        # Apply force password change if selected
        if ($result.Status -eq "Success" -and $forceChange) {
            try {
                Set-ADUser -Identity $user.DistinguishedName -ChangePasswordAtLogon $true
                Write-Host "  NOTE: User will be required to change password at next logon." -ForegroundColor Yellow
                $result.Message += " (Change at logon enforced)"
            }
            catch {
                Write-Host "  WARNING: Could not set 'change password at logon' flag: $_" -ForegroundColor Yellow
            }
        }
        
        $resultObject = [PSCustomObject]@{
            DisplayName = if ($result.DisplayName) { $result.DisplayName } else { $user.DisplayName }
            Username = if ($result.Username) { $result.Username } else { $user.SamAccountName }
            Status = if ($result.Status) { $result.Status } else { "Failed" }
            Message = if ($result.Message) { $result.Message } else { "Unknown error" }
            NewPassword = $resetPassword
            ForceChangeAtLogon = $forceChange
            OU = $ouInfo.FullPath
            OUName = $ouInfo.Name
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
    
    # Display detailed results
    if ($results.Count -gt 0) {
        Write-Host "`n--- Detailed Results ---" -ForegroundColor White
        $results | Format-Table DisplayName, Username, Status -AutoSize | Out-Host
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