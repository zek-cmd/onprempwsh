<#
.SYNOPSIS
    Creates and configures a shared folder for GPO software deployment with proper permissions.

.DESCRIPTION
    This script creates a folder, shares it, and sets appropriate Share and NTFS permissions
    for software deployment via Group Policy. It ensures Domain Computers have the necessary
    read access for computer-based installation.

.PARAMETER FolderPath
    The local path where the deployment folder will be created (e.g., D:\GPODeploy)

.PARAMETER ShareName
    The name of the network share (e.g., GPODeploy$)

.PARAMETER DomainName
    The NETBIOS domain name (e.g., CONTOSO). If not specified, tries to detect automatically.

.PARAMETER Interactive
    Use GUI dialogs for folder selection (default: $true)

.EXAMPLE
    .\Create-GPODeploymentShare.ps1 -Interactive

.EXAMPLE
    .\Create-GPODeploymentShare.ps1 -FolderPath "D:\SoftwareDeployment" -ShareName "SoftwareDeploy$"

.EXAMPLE
    .\Create-GPODeploymentShare.ps1 -FolderPath "C:\GPOPackages" -ShareName "Apps$" -DomainName "CONTOSO"

.NOTES
    Author: System Administrator
    Version: 2.0
    Requires: PowerShell 5.1+, Administrative privileges
#>

[CmdletBinding(DefaultParameterSetName='Interactive')]
param(
    [Parameter(Mandatory=$false, ParameterSetName='Manual')]
    [ValidateScript({Test-Path (Split-Path $_ -Parent)})]
    [string]$FolderPath,
    
    [Parameter(Mandatory=$false, ParameterSetName='Manual')]
    [ValidatePattern('^[a-zA-Z0-9_\-$]+$')]
    [string]$ShareName,
    
    [Parameter(Mandatory=$false)]
    [string]$DomainName,
    
    [Parameter(Mandatory=$false, ParameterSetName='Interactive')]
    [switch]$Interactive = $true
)

#region Initialization
# Require admin privileges
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator!" -ErrorAction Stop
}

# Load required assemblies for folder browser dialog
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Get domain name if not provided
if (-not $DomainName) {
    try {
        $DomainName = (Get-ADDomain).NetBIOSName
        Write-Host "Detected domain: $DomainName" -ForegroundColor Green
    }
    catch {
        try {
            $DomainName = $env:USERDOMAIN
            Write-Host "Using current domain: $DomainName" -ForegroundColor Yellow
        }
        catch {
            Write-Error "Cannot determine domain name. Please specify using -DomainName parameter." -ErrorAction Stop
        }
    }
}

# Validate domain exists
$domainCheck = $null
try {
    $domainCheck = Get-ADDomain -Identity $DomainName -ErrorAction Stop
}
catch {
    Write-Warning "Could not validate domain '$DomainName' in Active Directory. Continuing anyway..."
}
#endregion

#region GUI Functions
function Select-FolderDialog {
    param(
        [string]$Description = "Select folder for GPO software deployment share",
        [string]$InitialDirectory = "C:\"
    )
    
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Description
    $folderBrowser.SelectedPath = $InitialDirectory
    $folderBrowser.ShowNewFolderButton = $true
    
    $result = $folderBrowser.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    }
    else {
        return $null
    }
}

function Show-ConfirmationDialog {
    param(
        [string]$Message,
        [string]$Title = "Confirmation Required"
    )
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2
    )
    
    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Show-ShareNameDialog {
    param(
        [string]$FolderName
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Configure Share Name"
    $form.Size = New-Object System.Drawing.Size(450, 200)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    
    # Suggested share name
    $suggestedName = if ($FolderName) { 
        "$FolderName`$" 
    } else { 
        "SoftwareDeploy$" 
    }
    
    # Label
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(15, 15)
    $label.Size = New-Object System.Drawing.Size(400, 40)
    $label.Text = "Enter the share name (use `$ at the end for hidden share):"
    $form.Controls.Add($label)
    
    # TextBox for share name
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(15, 60)
    $textBox.Size = New-Object System.Drawing.Size(400, 20)
    $textBox.Text = $suggestedName
    $form.Controls.Add($textBox)
    
    # OK Button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(250, 100)
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton
    
    # Cancel Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(335, 100)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton
    
    # Show form
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $shareName = $textBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($shareName)) {
            Write-Warning "Share name cannot be empty. Using default: GPODeploy$"
            return "GPODeploy$"
        }
        # Validate share name format
        if ($shareName -match '^[a-zA-Z0-9_\-$]+$') {
            return $shareName
        }
        else {
            Write-Warning "Invalid share name format. Using default: GPODeploy$"
            return "GPODeploy$"
        }
    }
    else {
        return $null
    }
}

function Show-SummaryDialog {
    param(
        [string]$FolderPath,
        [string]$ShareName,
        [string]$DomainName,
        [string]$ServerName,
        [string]$UNCPath
    )
    
    $summaryText = @"
Configuration Summary:
======================

Domain:        $DomainName
Server:        $ServerName
Local Path:    $FolderPath
Share Name:    $ShareName
UNC Path:      $UNCPath

Share Permissions:
  • $DomainName\Domain Admins: Full Control
  • $DomainName\Domain Computers: Read
  • Authenticated Users: Read

NTFS Permissions:
  • SYSTEM: Full Control
  • Administrators: Full Control
  • $DomainName\Domain Admins: Full Control
  • $DomainName\Domain Computers: Read & Execute
  • Authenticated Users: Read & Execute

Do you want to proceed with this configuration?
"@
    
    return Show-ConfirmationDialog -Message $summaryText -Title "Confirm Share Configuration"
}
#endregion

#region Interactive Mode Functions
function Start-InteractiveMode {
    param(
        [string]$DomainName
    )
    
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "GPO Software Deployment Share Creator" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Interactive Mode" -ForegroundColor Cyan
    Write-Host ""
    
    # Step 1: Select folder
    Write-Host "Step 1: Select deployment folder..." -ForegroundColor Cyan
    $selectedPath = Select-FolderDialog -Description "Select or create folder for GPO software deployment"
    
    if (-not $selectedPath) {
        Write-Host "Folder selection cancelled. Exiting." -ForegroundColor Red
        return
    }
    Write-Host "  Selected: $selectedPath" -ForegroundColor Green
    
    # Step 2: Get share name
    Write-Host "`nStep 2: Configure share name..." -ForegroundColor Cyan
    $folderName = Split-Path $selectedPath -Leaf
    $shareName = Show-ShareNameDialog -FolderName $folderName
    
    if (-not $shareName) {
        Write-Host "Share name configuration cancelled. Exiting." -ForegroundColor Red
        return
    }
    Write-Host "  Share Name: $shareName" -ForegroundColor Green
    
    # Step 3: Show summary and confirm
    Write-Host "`nStep 3: Review configuration..." -ForegroundColor Cyan
    $serverName = $env:COMPUTERNAME
    $uncPath = "\\$serverName\$shareName"
    
    $confirmed = Show-SummaryDialog -FolderPath $selectedPath -ShareName $shareName `
                                    -DomainName $DomainName -ServerName $serverName -UNCPath $uncPath
    
    if (-not $confirmed) {
        Write-Host "Configuration cancelled by user." -ForegroundColor Red
        return
    }
    
    # Step 4: Execute configuration
    Write-Host "`nConfiguration confirmed. Proceeding with setup..." -ForegroundColor Green
    Configure-DeploymentShare -FolderPath $selectedPath -ShareName $shareName -DomainName $DomainName
}
#endregion

#region Configuration Functions
function Set-SharePermissions {
    param(
        [string]$ShareName,
        [string]$Domain
    )
    
    Write-Host "`nConfiguring Share Permissions..." -ForegroundColor Cyan
    
    # Remove existing share permissions (if any)
    try {
        $share = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
        if ($share) {
            $currentPerms = Get-SmbShareAccess -Name $ShareName
            foreach ($perm in $currentPerms) {
                Revoke-SmbShareAccess -Name $ShareName -AccountName $perm.AccountName -Force
                Write-Host "  Removed existing share permission for: $($perm.AccountName)" -ForegroundColor Gray
            }
        }
    }
    catch {
        Write-Warning "Could not clear existing share permissions: $_"
    }
    
    # Set share permissions
    $sharePerms = @(
        @{
            Account = "$Domain\Domain Admins"
            AccessRight = "Full"
            Description = "Domain Admins - Full Control"
        },
        @{
            Account = "$Domain\Domain Computers"
            AccessRight = "Read"
            Description = "Domain Computers - Read (Required for computer-based deployment)"
        },
        @{
            Account = "Authenticated Users"
            AccessRight = "Read"
            Description = "Authenticated Users - Read (For user-based deployment)"
        }
    )
    
    foreach ($perm in $sharePerms) {
        try {
            Grant-SmbShareAccess -Name $ShareName -AccountName $perm.Account -AccessRight $perm.AccessRight -Force
            Write-Host "  ✓ Granted $($perm.AccessRight) to $($perm.Account) - $($perm.Description)" -ForegroundColor Green
        }
        catch {
            Write-Warning "  ✗ Failed to grant share permission to $($perm.Account): $_"
        }
    }
}

function Set-NTFSPermissions {
    param(
        [string]$Path,
        [string]$Domain
    )
    
    Write-Host "`nConfiguring NTFS Permissions..." -ForegroundColor Cyan
    
    # Get current ACL
    $acl = Get-Acl -Path $Path
    
    # Remove inheritable permissions from parent and convert to explicit
    $acl.SetAccessRuleProtection($true, $false)
    
    # Clear existing explicit permissions
    $acl.Access | ForEach-Object {
        $acl.RemoveAccessRule($_)
    }
    
    # Set NTFS permissions
    $ntfsPerms = @(
        @{
            Account = "NT AUTHORITY\SYSTEM"
            Rights = "FullControl"
            Type = "Allow"
            Description = "SYSTEM - Full Control"
        },
        @{
            Account = "BUILTIN\Administrators"
            Rights = "FullControl"
            Type = "Allow"
            Description = "Local Administrators - Full Control"
        },
        @{
            Account = "$Domain\Domain Admins"
            Rights = "FullControl"
            Type = "Allow"
            Description = "Domain Admins - Full Control"
        },
        @{
            Account = "$Domain\Domain Computers"
            Rights = "ReadAndExecute"
            Type = "Allow"
            Description = "Domain Computers - Read & Execute (Required for computer-based deployment)"
        },
        @{
            Account = "Authenticated Users"
            Rights = "ReadAndExecute"
            Type = "Allow"
            Description = "Authenticated Users - Read & Execute (For user-based deployment)"
        }
    )
    
    foreach ($perm in $ntfsPerms) {
        try {
            $rights = [System.Security.AccessControl.FileSystemRights]$perm.Rights
            $inheritanceFlags = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit,ObjectInherit"
            $propagationFlags = [System.Security.AccessControl.PropagationFlags]::None
            $accessControlType = [System.Security.AccessControl.AccessControlType]::Allow
            
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $perm.Account,
                $rights,
                $inheritanceFlags,
                $propagationFlags,
                $accessControlType
            )
            
            $acl.AddAccessRule($rule)
            Write-Host "  ✓ Added NTFS $($perm.Rights) for $($perm.Account) - $($perm.Description)" -ForegroundColor Green
        }
        catch {
            Write-Warning "  ✗ Failed to set NTFS permission for $($perm.Account): $_"
        }
    }
    
    # Apply the ACL
    try {
        Set-Acl -Path $Path -AclObject $acl
        Write-Host "  ✓ NTFS permissions applied successfully" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to apply NTFS permissions: $_" -ErrorAction Stop
    }
}

function Write-DeploymentInstructions {
    param(
        [string]$ServerName,
        [string]$Share,
        [string]$Path
    )
    
    $uncPath = "\\$ServerName\$Share"
    
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "DEPLOYMENT INSTRUCTIONS" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "UNC Path for GPO: " -NoNewline
    Write-Host "$uncPath" -ForegroundColor Green
    Write-Host ""
    Write-Host "File Location: " -NoNewline
    Write-Host "$uncPath\" -ForegroundColor Green
    Write-Host ""
    Write-Host "Important Reminders:" -ForegroundColor White
    Write-Host "  1. Always use UNC paths in Group Policy (not drive letters)" -ForegroundColor Gray
    Write-Host "  2. Example: " -NoNewline -ForegroundColor Gray
    Write-Host "$uncPath\application.msi" -ForegroundColor Green
    Write-Host "  3. Computer-based deployment uses Domain Computers account" -ForegroundColor Gray
    Write-Host "  4. Subfolders inherit these permissions automatically" -ForegroundColor Gray
    Write-Host "  5. Test deployment with gpupdate /force and reboot" -ForegroundColor Gray
    Write-Host ""
    Write-Host "To verify permissions:" -ForegroundColor White
    Write-Host "  Get-SmbShareAccess -Name '$Share'" -ForegroundColor Cyan
    Write-Host "  Get-Acl -Path '$Path' | Format-List" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Yellow
}

function Configure-DeploymentShare {
    param(
        [string]$FolderPath,
        [string]$ShareName,
        [string]$DomainName
    )
    
    # Step 1: Create folder if it doesn't exist
    Write-Host "`nCreating folder structure..." -ForegroundColor Cyan
    try {
        if (-not (Test-Path -Path $FolderPath)) {
            New-Item -Path $FolderPath -ItemType Directory -Force | Out-Null
            Write-Host "  ✓ Created folder: $FolderPath" -ForegroundColor Green
        }
        else {
            Write-Host "  ✓ Folder already exists: $FolderPath" -ForegroundColor Green
            
            # Confirm if folder already exists with content
            $existingItems = Get-ChildItem -Path $FolderPath -ErrorAction SilentlyContinue
            if ($existingItems) {
                Write-Host "  ⚠ Warning: Folder contains existing items. Permissions will be overwritten." -ForegroundColor Yellow
                $continue = Show-ConfirmationDialog -Message "The folder '$FolderPath' already contains files/folders.`n`nExisting permissions will be overwritten.`n`nDo you want to continue?" -Title "Folder Already Contains Data"
                if (-not $continue) {
                    Write-Host "Operation cancelled by user." -ForegroundColor Red
                    exit
                }
            }
        }
    }
    catch {
        Write-Error "Failed to create folder: $_" -ErrorAction Stop
    }
    
    # Step 2: Create or update the share
    Write-Host "`nCreating/Updating share..." -ForegroundColor Cyan
    try {
        $existingShare = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
        if ($existingShare) {
            Write-Host "  Share already exists. Updating path..." -ForegroundColor Yellow
            
            # Confirm share update
            $continue = Show-ConfirmationDialog -Message "Share '$ShareName' already exists.`n`nDo you want to update it to point to '$FolderPath'?" -Title "Share Already Exists"
            if (-not $continue) {
                Write-Host "Operation cancelled by user." -ForegroundColor Red
                exit
            }
            
            Set-SmbShare -Name $ShareName -Path $FolderPath -Force
            Write-Host "  ✓ Updated share: $ShareName -> $FolderPath" -ForegroundColor Green
        }
        else {
            New-SmbShare -Name $ShareName -Path $FolderPath -Description "GPO Software Deployment Share" -FullAccess "$DomainName\Domain Admins"
            Write-Host "  ✓ Created share: $ShareName -> $FolderPath" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Failed to create/update share: $_" -ErrorAction Stop
    }
    
    # Step 3: Set share permissions
    Set-SharePermissions -ShareName $ShareName -Domain $DomainName
    
    # Step 4: Set NTFS permissions
    Set-NTFSPermissions -Path $FolderPath -Domain $DomainName
    
    # Step 5: Create example subfolder structure
    Write-Host "`nCreating example deployment structure..." -ForegroundColor Cyan
    $subFolders = @("Applications", "Scripts", "MST-Files", "Logs")
    foreach ($folder in $subFolders) {
        $subPath = Join-Path -Path $FolderPath -ChildPath $folder
        if (-not (Test-Path -Path $subPath)) {
            New-Item -Path $subPath -ItemType Directory -Force | Out-Null
            Write-Host "  ✓ Created: $subPath" -ForegroundColor Green
        }
    }
    
    # Step 6: Display deployment instructions
    $serverName = $env:COMPUTERNAME
    Write-DeploymentInstructions -ServerName $serverName -Share $ShareName -Path $FolderPath
    
    # Step 7: Final verification
    Write-Host "`nRunning final verification..." -ForegroundColor Cyan
    $verificationPassed = $true
    
    # Verify share exists
    $shareCheck = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
    if (-not $shareCheck) {
        Write-Host "  ✗ Share verification failed!" -ForegroundColor Red
        $verificationPassed = $false
    }
    else {
        Write-Host "  ✓ Share exists and accessible" -ForegroundColor Green
    }
    
    # Verify share permissions
    $sharePermCheck = Get-SmbShareAccess -Name $ShareName | Where-Object {$_.AccountName -eq "$DomainName\Domain Computers"}
    if (-not $sharePermCheck) {
        Write-Host "  ✗ Domain Computers share permission missing!" -ForegroundColor Red
        $verificationPassed = $false
    }
    else {
        Write-Host "  ✓ Domain Computers has share access" -ForegroundColor Green
    }
    
    # Verify NTFS permissions
    $aclCheck = Get-Acl -Path $FolderPath
    $domainComputersNTFS = $aclCheck.Access | Where-Object {$_.IdentityReference -eq "$DomainName\Domain Computers"}
    if (-not $domainComputersNTFS) {
        Write-Host "  ✗ Domain Computers NTFS permission missing!" -ForegroundColor Red
        $verificationPassed = $false
    }
    else {
        Write-Host "  ✓ Domain Computers has NTFS access" -ForegroundColor Green
    }
    
    if ($verificationPassed) {
        Write-Host "`n✓ Configuration completed successfully!" -ForegroundColor Green
        
        # Show success message
        [System.Windows.Forms.MessageBox]::Show(
            "Share configuration completed successfully!`n`nUNC Path: \\$serverName\$ShareName`n`nYou can now deploy software via GPO using this path.",
            "Configuration Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    else {
        Write-Host "`n⚠ Configuration completed with warnings. Check the output above." -ForegroundColor Yellow
        
        [System.Windows.Forms.MessageBox]::Show(
            "Share configuration completed with warnings.`n`nPlease check the PowerShell output for details.",
            "Configuration Complete with Warnings",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
    
    Write-Host "`nScript completed at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
}
#endregion

#region Main Execution
Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "GPO Software Deployment Share Creator" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Domain: " -NoNewline
Write-Host $DomainName -ForegroundColor Green

if ($Interactive -and (-not $FolderPath -or -not $ShareName)) {
    # Run in interactive mode with GUI dialogs
    Start-InteractiveMode -DomainName $DomainName
}
else {
    # Run in manual mode with provided parameters
    if (-not $FolderPath -or -not $ShareName) {
        Write-Error "When not using interactive mode, both -FolderPath and -ShareName are required." -ErrorAction Stop
    }
    
    Write-Host "Folder: " -NoNewline
    Write-Host $FolderPath -ForegroundColor Green
    Write-Host "Share:  " -NoNewline
    Write-Host $ShareName -ForegroundColor Green
    
    # Show confirmation even in manual mode
    $serverName = $env:COMPUTERNAME
    $uncPath = "\\$serverName\$ShareName"
    $confirmed = Show-SummaryDialog -FolderPath $FolderPath -ShareName $ShareName `
                                    -DomainName $DomainName -ServerName $serverName -UNCPath $uncPath
    
    if (-not $confirmed) {
        Write-Host "Configuration cancelled by user." -ForegroundColor Red
        exit
    }
    
    Configure-DeploymentShare -FolderPath $FolderPath -ShareName $ShareName -DomainName $DomainName
}
#endregion