# Active Directory User Move Script (CSV-Based, FullName Only)
# Move users from a source OU to a target OU based on FullName in CSV
# Usernames are generated from FullName using initials + last name pattern
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

# Function to remove special characters and normalize text for username
function Normalize-ForUsername {
    param([string]$Text)
    
    $normalized = $Text.ToLower()
    $normalized = $normalized.Normalize([Text.NormalizationForm]::FormC)
    
    $charMap = @{}
    
    $charMap[[char]0x00E1] = 'a'
    $charMap[[char]0x00E0] = 'a'
    $charMap[[char]0x00E4] = 'a'
    $charMap[[char]0x00E2] = 'a'
    $charMap[[char]0x00E3] = 'a'
    $charMap[[char]0x00E5] = 'a'
    
    $charMap[[char]0x00E9] = 'e'
    $charMap[[char]0x00E8] = 'e'
    $charMap[[char]0x00EB] = 'e'
    $charMap[[char]0x00EA] = 'e'
    
    $charMap[[char]0x00ED] = 'i'
    $charMap[[char]0x00EC] = 'i'
    $charMap[[char]0x00EF] = 'i'
    $charMap[[char]0x00EE] = 'i'
    
    $charMap[[char]0x00F3] = 'o'
    $charMap[[char]0x00F2] = 'o'
    $charMap[[char]0x00F6] = 'o'
    $charMap[[char]0x00F4] = 'o'
    $charMap[[char]0x00F5] = 'o'
    $charMap[[char]0x00F8] = 'o'
    
    $charMap[[char]0x00FA] = 'u'
    $charMap[[char]0x00F9] = 'u'
    $charMap[[char]0x00FC] = 'u'
    $charMap[[char]0x00FB] = 'u'
    
    $charMap[[char]0x00F1] = 'n'
    $charMap[[char]0x00E7] = 'c'
    $charMap[[char]0x00FD] = 'y'
    $charMap[[char]0x00FF] = 'y'
    
    $charMap[[char]0x00DF] = 'ss'
    $charMap[[char]0x00E6] = 'ae'
    $charMap[[char]0x0153] = 'oe'
    
    $result = ""
    for ($i = 0; $i -lt $normalized.Length; $i++) {
        $char = $normalized[$i]
        
        if ($charMap.ContainsKey($char)) {
            $result += $charMap[$char]
        }
        elseif ([char]::IsLetterOrDigit($char)) {
            $result += $char
        }
    }
    
    return $result
}

# Function to generate username from full name
function Generate-Username {
    param([string]$FullName)
    
    try {
        $NameParts = Parse-NameParts $FullName
        
        $firstNames = $NameParts.FirstName -split '\s+'
        $firstInitials = ""
        foreach ($name in $firstNames) {
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $initial = Normalize-ForUsername $name.Substring(0, 1)
                $firstInitials += $initial
            }
        }
        
        $middleInitials = ""
        if (-not [string]::IsNullOrWhiteSpace($NameParts.MiddleName)) {
            $middleNames = $NameParts.MiddleName -split '\s+'
            foreach ($name in $middleNames) {
                if (-not [string]::IsNullOrWhiteSpace($name)) {
                    $cleanName = $name -replace '\.', ''
                    $initial = Normalize-ForUsername $cleanName.Substring(0, 1)
                    $middleInitials += $initial
                }
            }
        }
        
        $lastName = Normalize-ForUsername $NameParts.LastName
        $username = $firstInitials + $middleInitials + $lastName
        
        if ([string]::IsNullOrWhiteSpace($username)) {
            return $null
        }
        
        return $username.ToLower()
    }
    catch {
        return $null
    }
}

# Function to extract name parts for AD attributes
function Parse-NameParts {
    param([string]$FullName)
    
    $parsingName = $FullName -replace '\s+', ' '
    $parsingName = $parsingName.Trim()
    
    if ([string]::IsNullOrWhiteSpace($parsingName)) {
        return @{
            FirstName = ""
            LastName = ""
            MiddleName = ""
            DisplayName = ""
        }
    }
    
    $parts = $parsingName -split '\s+'
    
    $result = @{
        FirstName = ""
        LastName = ""
        MiddleName = ""
        DisplayName = $parsingName
    }
    
    if ($parts.Count -eq 1) {
        $result.FirstName = $parts[0]
        $result.LastName = $parts[0]
    }
    elseif ($parts.Count -eq 2) {
        $result.FirstName = $parts[0]
        $result.LastName = $parts[1]
    }
    else {
        $firstNames = @($parts[0])
        $result.LastName = $parts[-1]
        
        $middleParts = @()
        
        for ($i = 1; $i -lt ($parts.Count - 1); $i++) {
            $part = $parts[$i]
            
            if ($part -match '^[A-Za-z]{1,3}\.?$') {
                $middleParts += $part
            }
            else {
                $firstNames += $part
            }
        }
        
        $result.FirstName = $firstNames -join ' '
        
        if ($middleParts.Count -gt 0) {
            $result.MiddleName = $middleParts -join ' '
        }
    }
    
    $result.FirstName = $result.FirstName.Trim()
    $result.LastName = $result.LastName.Trim()
    $result.MiddleName = $result.MiddleName.Trim()
    
    return $result
}

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

# Function to read names from CSV file (FullName column only)
function Read-NamesFromCSV {
    param([string]$FilePath)
    
    if (-not (Test-Path -Path $FilePath -PathType Leaf)) {
        Write-Error "File not found: $FilePath"
        return @()
    }
    
    Write-Host "Reading names from CSV file..." -ForegroundColor White
    
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
        
        $names = @()
        foreach ($row in $csvData) {
            $fullName = $row.FullName.Trim()
            if (-not [string]::IsNullOrWhiteSpace($fullName)) {
                $names += $fullName
            }
        }
        
        Write-Host "Successfully read $($names.Count) names from CSV file." -ForegroundColor Green
        return $names
    }
    catch {
        Write-Error "Error reading CSV file: $_"
        return @()
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

# Function to move AD user to target OU (looked up by generated username)
function Move-ADUserToOU {
    param(
        [string]$FullName,
        [string]$TargetOUPath
    )
    
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        Write-Error "Active Directory module is not available. Please install RSAT tools."
        return @{
            Status = "Failed"
            Message = "AD module not available"
            Username = ""
            CurrentOU = ""
        }
    }
    
    $Username = Generate-Username $FullName
    
    if (-not $Username) {
        Write-Error "Failed to generate username for: $FullName"
        return @{
            Status = "Failed"
            Message = "Username generation failed"
            Username = ""
            CurrentOU = ""
        }
    }
    
    Write-Host "`nLooking up user: $FullName" -ForegroundColor White
    Write-Host "  Generated username: $Username" -ForegroundColor Gray
    
    $adUser = $null
    try {
        $adUser = Get-ADUser -Identity $Username -Properties DistinguishedName, CanonicalName, Name, SamAccountName
    }
    catch {
        # User not found
    }
    
    if (-not $adUser) {
        Write-Host "  User not found by username '$Username'. Trying display name fallback..." -ForegroundColor Yellow
        
        try {
            $adUser = Get-ADUser -Filter "DisplayName -eq '$FullName'" -Properties DistinguishedName, CanonicalName, Name, SamAccountName
        }
        catch {
            # Still not found
        }
    }
    
    if (-not $adUser) {
        Write-Host "  ERROR: User '$FullName' (username: $Username) not found in Active Directory." -ForegroundColor Red
        return @{
            Status = "Failed"
            Message = "User not found in AD"
            Username = $Username
            CurrentOU = ""
        }
    }
    
    $currentOU = $adUser.DistinguishedName -replace '^CN=.*?,(?=OU=|CN=)', ''
    
    Write-Host "  Found user: $($adUser.Name)" -ForegroundColor Green
    Write-Host "  Current Location: $currentOU" -ForegroundColor Gray
    Write-Host "  Target Location: $TargetOUPath" -ForegroundColor Yellow
    
    if ($currentOU -eq $TargetOUPath) {
        Write-Host "  SKIPPED: User is already in the target OU." -ForegroundColor Yellow
        return @{
            Status = "Skipped"
            Message = "Already in target OU"
            Username = $adUser.SamAccountName
            CurrentOU = $currentOU
        }
    }
    
    try {
        Move-ADObject -Identity $adUser.DistinguishedName -TargetPath $TargetOUPath
        Write-Host "  SUCCESS: Moved to target OU." -ForegroundColor Green
        
        return @{
            Status = "Success"
            Message = "Moved successfully"
            Username = $adUser.SamAccountName
            CurrentOU = $currentOU
        }
    }
    catch {
        Write-Host "  ERROR: Failed to move user: $_" -ForegroundColor Red
        
        return @{
            Status = "Failed"
            Message = "Move operation failed: $_"
            Username = $adUser.SamAccountName
            CurrentOU = $currentOU
        }
    }
}

# Function to display summary before execution
function Show-Summary {
    param(
        [array]$Names,
        [hashtable]$OUInfo,
        [string]$NamesFile
    )
    
    Write-Host "`n=== MOVE SUMMARY ===" -ForegroundColor Cyan
    Write-Host "Names file: $NamesFile" -ForegroundColor Yellow
    Write-Host "Number of users to move: $($Names.Count)" -ForegroundColor Yellow
    Write-Host "Target OU: $($OUInfo.Name)" -ForegroundColor Yellow
    Write-Host "  OU Path: $($OUInfo.FullPath)" -ForegroundColor Gray
    Write-Host "Domain: $Domain" -ForegroundColor Yellow
    Write-Host "Lookup Method: Generated username from FullName" -ForegroundColor Yellow
    
    Write-Host "`n--- Name to Username Mapping ---" -ForegroundColor White
    
    $maxNameLength = 0
    foreach ($name in $Names) {
        if ($name.Length -gt $maxNameLength) {
            $maxNameLength = $name.Length
        }
    }
    $maxNameLength += 5
    
    foreach ($name in $Names) {
        $username = Generate-Username $name
        $paddedName = $name.PadRight($maxNameLength)
        Write-Host "$paddedName -> $username" -ForegroundColor Gray
    }
    
    Write-Host "`n=================================" -ForegroundColor Cyan
}

# Main execution
function Main {
    Clear-Host
    
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "     ACTIVE DIRECTORY USER MOVE TOOL" -ForegroundColor Cyan
    Write-Host "     (CSV-Based: FullName column only)" -ForegroundColor Cyan
    Write-Host "     (Username generated from FullName)" -ForegroundColor Cyan
    Write-Host "     (Graphical OU Browser)" -ForegroundColor Cyan
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
    
    # Step 1: Get names CSV file
    Write-Host "`nStep 1: Select names.csv file" -ForegroundColor White
    $namesFile = Get-NamesFile
    if (-not $namesFile) {
        Write-Host "Operation cancelled by user." -ForegroundColor Red
        return
    }
    
    # Step 2: Read names from CSV file
    $names = Read-NamesFromCSV -FilePath $namesFile
    if ($names.Count -eq 0) {
        Write-Host "No valid names found in the file." -ForegroundColor Red
        Write-Host "Make sure your CSV has a 'FullName' column header." -ForegroundColor Yellow
        return
    }
    
    # Step 3: Get target OU using tree browser
    $ouInfo = Get-TargetOU
    if (-not $ouInfo) {
        Write-Host "OU selection failed or cancelled." -ForegroundColor Red
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($ouInfo.FullPath)) {
        Write-Host "ERROR: OU path is empty. Cannot move users." -ForegroundColor Red
        return
    }
    
    # Step 4: Show summary
    Show-Summary -Names $names -OUInfo $ouInfo -NamesFile $namesFile
    
    # Step 5: Ask for confirmation
    Write-Host "`n=== CONFIRMATION REQUIRED ===" -ForegroundColor Red
    Write-Host "This will move $($names.Count) user accounts to:" -ForegroundColor Yellow
    Write-Host "  Target OU: $($ouInfo.Name)" -ForegroundColor Yellow
    Write-Host "  OU Path: $($ouInfo.FullPath)" -ForegroundColor Gray
    Write-Host "  Domain: $Domain" -ForegroundColor Yellow
    Write-Host "  Lookup: Generated username from FullName" -ForegroundColor Yellow
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
    
    foreach ($fullName in $names) {
        Write-Host "`n" + ("="*50) -ForegroundColor DarkGray
        Write-Host "Processing: $fullName" -ForegroundColor White
        
        $result = Move-ADUserToOU -FullName $fullName -TargetOUPath $ouInfo.FullPath
        
        $resultObject = [PSCustomObject]@{
            FullName = $fullName
            Username = if ($result.Username) { $result.Username } else { "" }
            Status = if ($result.Status) { $result.Status } else { "Failed" }
            Message = if ($result.Message) { $result.Message } else { "Unknown error" }
            PreviousOU = if ($result.CurrentOU) { $result.CurrentOU } else { "" }
            TargetOU = $ouInfo.FullPath
            TargetOUName = $ouInfo.Name
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
        $results | Format-Table FullName, Username, Status -AutoSize | Out-Host
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