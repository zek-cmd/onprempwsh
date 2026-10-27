# Active Directory User Creation Script
# Enhanced version with OU tree browser and proper special character handling
# Updated: Reads names from CSV file (FullName and StudNo columns)
# Updated: Uses StudNo as username instead of generating from full name
# Updated: Existing users are skipped entirely (no changes made)
# Script compatibility verified for PowerShell 5.1+

# Import required assemblies for dialog boxes
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.DirectoryServices.Protocols -ErrorAction SilentlyContinue

# Configuration
$InitialPassword = "P@ssw0rd"

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

# Get domain configuration
$domainConfig = Get-CurrentDomainInfo

if (-not $domainConfig) {
    Write-Host "Press Enter to exit..." -ForegroundColor Gray
    Read-Host
    exit 1
}

# Set domain variables
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

# Function to generate username from full name (kept for fallback/compatibility)
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
        [string]$DefaultFileName = "AD_User_Creation_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    )
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveFileDialog.Title = "Save Results CSV Report"
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

# Function to read names from CSV file (FullName + StudNo columns)
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
        
        # Check for required columns
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
            
            if (-not [string]::IsNullOrWhiteSpace($fullName) -and -not [string]::IsNullOrWhiteSpace($studNo)) {
                $entries += [PSCustomObject]@{
                    FullName = $fullName
                    StudNo   = $studNo
                }
            }
            elseif (-not [string]::IsNullOrWhiteSpace($fullName)) {
                Write-Warning "Row has FullName '$fullName' but empty StudNo. Skipping."
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
    Write-Host "`n=== OU Selection ===" -ForegroundColor Cyan
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
    
    Write-Host "Selected OU: $($selectedOU.Name)" -ForegroundColor Green
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

# Function to create AD user using StudNo as username
function Create-ADUserFromName {
    param(
        [string]$FullName,
        [string]$StudNo,
        [string]$TargetOUPath
    )
    
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        Write-Error "Active Directory module is not available. Please install RSAT tools."
        return @{
            Status = "Failed"
            Message = "AD module not available"
            Username = ""
        }
    }
    
    # Use student number as the username directly
    $Username = $StudNo.Trim()
    
    if ([string]::IsNullOrWhiteSpace($Username)) {
        Write-Error "Student number is empty for: $FullName"
        return @{
            Status = "Failed"
            Message = "Student number is empty"
            Username = ""
        }
    }
    
    $NameParts = Parse-NameParts $FullName
    $DisplayName = $FullName
    
    # Check if user already exists
    $ExistingUser = Get-ADUser -Filter "SamAccountName -eq '$Username'" -ErrorAction SilentlyContinue
    
    # Handle existing user - SKIP entirely
    if ($ExistingUser) {
        Write-Host "`nUser '$Username' already exists in Active Directory" -ForegroundColor Yellow
        Write-Host "  Full Name: $($ExistingUser.Name)" -ForegroundColor Gray
        Write-Host "  Status: Skipped (user already exists, no changes made)" -ForegroundColor Yellow
        
        return @{
            Status = "Skipped"
            Message = "User already exists - no changes made"
            Username = $Username
            Password = $InitialPassword
            UPN = $ExistingUser.UserPrincipalName
            FirstName = $NameParts.FirstName
            LastName = $NameParts.LastName
            DisplayName = $DisplayName
        }
    }
    
    # User does not exist - create new account
    $SecurePassword = ConvertTo-SecureString $InitialPassword -AsPlainText -Force
    $UserPrincipalName = "$Username@$Domain"
    
    Write-Host "`nCreating account for: $FullName" -ForegroundColor Cyan
    Write-Host "  StudNo / Username: $Username" -ForegroundColor Green
    Write-Host "  Initial Password: $InitialPassword" -ForegroundColor Yellow
    Write-Host "  User Principal Name: $UserPrincipalName" -ForegroundColor Green
    Write-Host "  Display Name: $DisplayName" -ForegroundColor Green
    Write-Host "  First Name: $($NameParts.FirstName)" -ForegroundColor Gray
    Write-Host "  Last Name: $($NameParts.LastName)" -ForegroundColor Gray
    if ($NameParts.MiddleName) {
        Write-Host "  Middle Name: $($NameParts.MiddleName)" -ForegroundColor Gray
    }
    Write-Host "  Target OU: $TargetOUPath" -ForegroundColor Gray
    
    $NewUserParams = @{
        SamAccountName        = $Username
        UserPrincipalName     = $UserPrincipalName
        Name                  = "$($NameParts.FirstName) $($NameParts.LastName)"
        GivenName             = $NameParts.FirstName
        Surname               = $NameParts.LastName
        DisplayName           = $DisplayName
        AccountPassword       = $SecurePassword
        Enabled               = $true
        PasswordNeverExpires  = $false
        ChangePasswordAtLogon = $true
        Path                  = $TargetOUPath
    }
    
    if (-not [string]::IsNullOrWhiteSpace($NameParts.MiddleName)) {
        $NewUserParams['OtherName'] = $NameParts.MiddleName
    }
    
    try {
        New-ADUser @NewUserParams
        Write-Host "  Status: Successfully created" -ForegroundColor Green
        
        return @{
            Status = "Success"
            Message = "Account created successfully"
            Username = $Username
            Password = $InitialPassword
            UPN = $UserPrincipalName
            FirstName = $NameParts.FirstName
            LastName = $NameParts.LastName
            DisplayName = $DisplayName
        }
    }
    catch {
        Write-Host "  ERROR: Failed to create user: $_" -ForegroundColor Red
        
        return @{
            Status = "Failed"
            Message = "Account creation failed: $_"
            Username = $Username
        }
    }
}

# Function to display summary before execution
function Show-Summary {
    param(
        [array]$Entries,
        [hashtable]$OUInfo,
        [string]$NamesFile
    )
    
    Write-Host "`n=== EXECUTION SUMMARY ===" -ForegroundColor Cyan
    Write-Host "Names file: $NamesFile" -ForegroundColor Yellow
    Write-Host "Number of entries to process: $($Entries.Count)" -ForegroundColor Yellow
    Write-Host "Target OU: $($OUInfo.Name)" -ForegroundColor Yellow
    Write-Host "  OU Path: $($OUInfo.FullPath)" -ForegroundColor Gray
    Write-Host "Domain: $Domain" -ForegroundColor Yellow
    Write-Host "Initial Password: $InitialPassword" -ForegroundColor Yellow
    Write-Host "Username Source: StudNo column" -ForegroundColor Yellow
    Write-Host "NOTE: Existing users will be SKIPPED entirely (no changes made)" -ForegroundColor Yellow
    
   Write-Host "`n--- FullName -> StudNo (Username) ---" -ForegroundColor White

# Find the longest FullName for padding
$maxNameLength = 0
foreach ($entry in $Entries) {
    if ($entry.FullName.Length -gt $maxNameLength) {
        $maxNameLength = $entry.FullName.Length
    }
}

# Add some padding
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
    Write-Host "     ACTIVE DIRECTORY USER CREATION TOOL" -ForegroundColor Cyan
    Write-Host "     (Enhanced Special Character Support)" -ForegroundColor Cyan
    Write-Host "     (Graphical OU Browser)" -ForegroundColor Cyan
    Write-Host "     (CSV Input: FullName + StudNo columns)" -ForegroundColor Cyan
    Write-Host "     (Username = StudNo)" -ForegroundColor Cyan
    Write-Host "     (Existing Users: Skip Entirely)" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Checking prerequisites..." -ForegroundColor White
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        Write-Host "WARNING: Active Directory module not found." -ForegroundColor Red
        Write-Host "Please install Remote Server Administration Tools (RSAT)." -ForegroundColor Red
        Write-Host "Script will continue but cannot create users without AD module." -ForegroundColor Yellow
        $continue = Read-Host "Continue anyway? (yes/no)"
        if ($continue -ne 'yes') {
            Write-Host "Script terminated by user." -ForegroundColor Yellow
            return
        }
    }
    
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
    
    # Step 3: Get target OU using tree browser
    $ouInfo = Get-TargetOU
    if (-not $ouInfo) {
        Write-Host "OU selection failed or cancelled." -ForegroundColor Red
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($ouInfo.FullPath)) {
        Write-Host "ERROR: OU path is empty. Cannot create users." -ForegroundColor Red
        return
    }
    
    # Step 4: Show summary
    Show-Summary -Entries $entries -OUInfo $ouInfo -NamesFile $namesFile
    
    # Step 5: Ask for confirmation
    Write-Host "`n=== CONFIRMATION REQUIRED ===" -ForegroundColor Red
    Write-Host "This will process $($entries.Count) user accounts in:" -ForegroundColor Yellow
    Write-Host "  OU: $($ouInfo.Name)" -ForegroundColor Yellow
    Write-Host "  OU Path: $($ouInfo.FullPath)" -ForegroundColor Gray
    Write-Host "  Domain: $Domain" -ForegroundColor Yellow
    Write-Host "  Username source: StudNo column" -ForegroundColor Cyan
    Write-Host "  New users: Created with password $InitialPassword" -ForegroundColor Green
    Write-Host "  Existing users: SKIPPED entirely (no changes made)" -ForegroundColor Yellow
    $confirmation = Read-Host "`nType 'YES' to proceed, anything else to cancel"
    
    if ($confirmation -ne "YES") {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        return
    }
    
    # Step 6: Create accounts
    Write-Host "`nStarting account processing..." -ForegroundColor Green
    
    $results = @()
    $successCount = 0
    $skipCount = 0
    $failCount = 0
    
    foreach ($entry in $entries) {
        Write-Host "`n" + ("="*50) -ForegroundColor DarkGray
        Write-Host "Processing: $($entry.FullName) (StudNo: $($entry.StudNo))" -ForegroundColor White
        
        $result = Create-ADUserFromName -FullName $entry.FullName -StudNo $entry.StudNo -TargetOUPath $ouInfo.FullPath
        
        $resultObject = [PSCustomObject]@{
            FullName = $entry.FullName
            StudNo   = $entry.StudNo
            Username = if ($result.Username) { $result.Username } else { "" }
            Status   = if ($result.Status) { $result.Status } else { "Failed" }
            Message  = if ($result.Message) { $result.Message } else { "Unknown error" }
            Password = $InitialPassword
            UPN      = if ($result.UPN) { $result.UPN } else { "" }
            OU       = $ouInfo.FullPath
            OUName   = $ouInfo.Name
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
    Write-Host "PROCESSING COMPLETE" -ForegroundColor Cyan
    Write-Host ("="*60) -ForegroundColor Cyan
    
    if ($successCount -gt 0) {
        Write-Host "Successfully created: $successCount accounts" -ForegroundColor Green
    }
    
    if ($skipCount -gt 0) {
        Write-Host "Skipped (existing users, no changes made): $skipCount accounts" -ForegroundColor Yellow
    }
    
    if ($failCount -gt 0) {
        Write-Host "Failed: $failCount accounts" -ForegroundColor Red
    }
    
    if ($results.Count -gt 0) {
        Write-Host "`n--- Detailed Results ---" -ForegroundColor White
        $results | Format-Table FullName, StudNo, Status -AutoSize | Out-Host
    }
    
    # Step 8: Let user choose where to save the results CSV
    Write-Host "`n" + ("-"*60) -ForegroundColor DarkGray
    Write-Host "Step 8: Save Results Report" -ForegroundColor White
    
    # Generate default filename with timestamp
    $defaultFileName = "AD_User_Creation_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
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
        Write-Host "`n--- Example New User Logins ---" -ForegroundColor White
        $successfulUsers = $results | Where-Object { $_.Status -eq "Success" } | Select-Object -First 5
        foreach ($user in $successfulUsers) {
            Write-Host "  Username: $($user.Username) | Password: $InitialPassword" -ForegroundColor Gray
        }
        
        if ($successfulUsers.Count -lt $successCount) {
            Write-Host "  ... and $($successCount - $successfulUsers.Count) more new users" -ForegroundColor Gray
        }
    }
    
    if ($skipCount -gt 0) {
        Write-Host "`n--- Existing Users (Skipped) ---" -ForegroundColor White
        $skippedUsers = $results | Where-Object { $_.Status -eq "Skipped" } | Select-Object -First 5
        foreach ($user in $skippedUsers) {
            Write-Host "  Username: $($user.Username) | Status: Skipped (no changes)" -ForegroundColor Gray
        }
        
        if ($skippedUsers.Count -lt $skipCount) {
            Write-Host "  ... and $($skipCount - $skippedUsers.Count) more existing users" -ForegroundColor Gray
        }
    }
    
    Write-Host "`nScript execution completed." -ForegroundColor Green
    Write-Host "Press Enter to exit..." -ForegroundColor Gray
    Read-Host
}

# Run the script
Main