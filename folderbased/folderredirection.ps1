<#
.SYNOPSIS
    Builds isolated per-user home folders on a file server: root folder with strict NTFS,
    hidden SMB share with Access-Based Enumeration (ABE), a Group Policy Object with
    Folder Redirection (Desktop / Documents / Downloads) and a mapped drive, all built
    on dynamic %USERNAME% paths.

.DESCRIPTION
    Steps:
      1. Create (or reuse) the security group that scopes the policy.
      2. Create the local root folder and lock its NTFS permissions down.
      3. Create (or reuse) the hidden SMB share, with ABE on.
      4. Create (or reuse) the GPO, set security filtering, link it to the target OU.
      5. Write fdeploy.ini (Folder Redirection) and Drives.xml (GPP Drive Maps) into
         SYSVOL, register both client-side extensions in AD, bump the GPO user version
         and sync GPT.INI.
      6. Wait for SYSVOL to reach the other domain controllers and print how to verify.

    Every step is idempotent: run it twice and the second run changes nothing.

.NOTES
    Requires: Domain Admin rights, RSAT (ActiveDirectory + GroupPolicy modules).
    Run it as Administrator on the file server, or on a management PC that can reach
    the server's local disk path and SYSVOL over the network.
#>

#Requires -RunAsAdministrator
#Requires -Modules ActiveDirectory,GroupPolicy

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================================
# 1. MANUAL CONFIGURATION - EDIT THESE TO MATCH YOUR ENVIRONMENT
# ============================================================================
$FileServerName        = ""                                      # Server holding the share. "" = the PC the script runs on
$LocalRootFolderPath   = "D:\UserRedirection"                   # Local drive + root folder on the server
$ShareName             = "UserRedirection$"                     # SMB share name ($ = hidden share)

# Subfolder names created inside each user's %USERNAME% directory
$DesktopFolderName     = "Desktop"
$DocumentsFolderName   = "Documents"
$DownloadsFolderName   = "Downloads"

# Active Directory & Group Policy settings
$DomainFQDN            = ""                                      # e.g. contoso.com - leave "" to auto-detect
$DomainNetBIOS         = ""                                      # e.g. CONTOSO - leave "" to auto-detect
$TargetGroup           = "SG-FolderRedirection"                 # Security group that gets the policy
$GPOName               = "User - Folder Redirection & Z Drive"  # Existing or new GPO name
$TargetOU              = ""                                      # e.g. OU=Users,DC=contoso,DC=com - "" = default Users container

# Drive mapping settings
$DriveLetter           = "Z"                                    # Drive letter, colon optional
$DriveLabel            = "Personal Network Drive"               # Label shown in File Explorer

# Folder redirection behaviour, written in the Flags= value of fdeploy.ini.
# 0x00001025 = create a folder per user, move existing contents into it, and do NOT
# grant the user exclusive rights - so Domain Admins / SYSTEM keep access.
$RedirectionFlags      = "0x00001025"

# Keep Downloads redirection on? Set to $false if your clients ignore that GUID.
$RedirectDownloads     = $true

# How long to wait for SYSVOL to copy to other domain controllers (0 = do not wait)
$SysvolWaitSeconds     = 90

# ============================================================================
# KNOWN FOLDER GUIDs (do not edit unless Microsoft changes them)
# ============================================================================
$DesktopGuid           = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
$DocumentsGuid         = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"
$DownloadsGuid         = "{374DE290-123F-4565-9164-39C4925E467B}"

# Client-side extensions that must be registered on the GPO in Active Directory
$FolderRedirectionCSE  = "[{25537BA6-77A8-11D2-9B6C-0000F8080861}{88E729D6-BDC1-11D1-BD2A-00C04FB9603F}]"
$DriveMapsCSE          = "[{5794DAFD-BE60-433f-88A2-1A31939AC01F}{2EA1A81B-48E5-45E9-8BB7-A6E3AC170006}]"

# Well-known SIDs - used instead of localized account names so non-English Windows works
$SidLocalAdmins        = "S-1-5-32-544"
$SidSystem             = "S-1-5-18"
$SidAuthenticatedUsers = "S-1-5-11"
$SidCreatorOwner       = "S-1-3-0"
$SidEveryone           = "S-1-1-0"

# ============================================================================
# HELPERS
# ============================================================================
function Write-Step {
    param([string]$Text)
    Write-Host "`n[+] $Text" -ForegroundColor Cyan
}

function Write-Ok {
    param([string]$Text)
    Write-Host "    OK  $Text" -ForegroundColor Green
}

function Write-Info {
    param([string]$Text)
    Write-Host "        $Text" -ForegroundColor Gray
}

function Write-Warn {
    param([string]$Text)
    Write-Host "    !!  $Text" -ForegroundColor Yellow
}

function Get-SidName {
    param([string]$Sid)
    try {
        return (New-Object System.Security.Principal.SecurityIdentifier($Sid)).Translate(
            [System.Security.Principal.NTAccount]).Value
    }
    catch {
        return $Sid
    }
}

$StartTime = Get-Date

try {
    # ========================================================================
    # 0. RESOLVE DOMAIN / PATHS
    # ========================================================================
    try {
        $AdDomain = Get-ADDomain -ErrorAction Stop
    }
    catch {
        throw "Cannot read the Active Directory domain from this machine. Run the script as a Domain Admin on a domain-joined server, or fill in the domain variables at the top. Reason: $($_.Exception.Message)"
    }

    $DomainDN = $AdDomain.DistinguishedName
    if ([string]::IsNullOrWhiteSpace($DomainFQDN))    { $DomainFQDN    = $AdDomain.DNSRoot }
    if ([string]::IsNullOrWhiteSpace($DomainNetBIOS)) { $DomainNetBIOS = $AdDomain.NetBIOSName }
    if ([string]::IsNullOrWhiteSpace($TargetOU))      { $TargetOU      = $AdDomain.UsersContainer }

    # Which server hosts the share? Auto-detect only makes sense when the script runs
    # on the file server itself - a domain controller has no business hosting this.
    $FileServerIsAuto = [string]::IsNullOrWhiteSpace($FileServerName)
    $LocalIsDC        = $null
    if ($FileServerIsAuto) {
        $FileServerName = $env:COMPUTERNAME
        $LocalIsDC = $AdDomain.DomainControllers | Where-Object { $_.Name -like "$($env:COMPUTERNAME)*" }
        if ($LocalIsDC) {
            Write-Warn "This PC is a domain controller, so '$FileServerName' is probably NOT the file server."
            Write-Warn "Set `$FileServerName (and `$LocalRootFolderPath) at the top to your real file server."
        }
    }
    $FileServerName = $FileServerName.Trim()

    Write-Info "Detected: domain '$DomainFQDN' (NetBIOS '$DomainNetBIOS')"
    Write-Info "Detected: OU '$TargetOU'"
    Write-Info "Detected: file server '$FileServerName'$(if ($FileServerIsAuto -and -not $LocalIsDC) { ' (this PC)' } else { '' })"

    $DriveLetter = $DriveLetter.Trim().TrimEnd(':').ToUpper()
    if ($DriveLetter -notmatch '^[A-Z]$') {
        throw "`$DriveLetter must be a single drive letter such as 'Z' (got '$DriveLetter')."
    }

    foreach ($Name in @($DesktopFolderName, $DocumentsFolderName, $DownloadsFolderName)) {
        if ([string]::IsNullOrWhiteSpace($Name)) { throw "Folder names cannot be empty." }
        if ($Name -match '[\\/:*?"<>|]') {
            throw "Folder name '$Name' contains a character Windows does not allow in a folder name."
        }
    }

    $UncRootPath      = "\\$FileServerName\$ShareName"
    $UncUserBasePath  = "$UncRootPath\%USERNAME%"
    $UncDesktopPath   = "$UncUserBasePath\$DesktopFolderName"
    $UncDocumentsPath = "$UncUserBasePath\$DocumentsFolderName"
    $UncDownloadsPath = "$UncUserBasePath\$DownloadsFolderName"

    Write-Host "=== Folder Redirection + Drive Map Deployment ===" -ForegroundColor White
    Write-Host "Domain          : $DomainFQDN ($DomainNetBIOS)"
    Write-Host "Share root      : $UncRootPath"
    Write-Host "Per-user root   : $UncUserBasePath"
    Write-Host "Desktop target  : $UncDesktopPath"
    Write-Host "Documents target: $UncDocumentsPath"
    if ($RedirectDownloads) { Write-Host "Downloads target: $UncDownloadsPath" }

    # ========================================================================
    # 1. SECURITY GROUP
    # ========================================================================
    Write-Step "1/6  Security group '$TargetGroup'"
    $Group = Get-ADGroup -Filter { Name -eq $TargetGroup } -ErrorAction SilentlyContinue
    if (-not $Group) {
        $Group = New-ADGroup -Name $TargetGroup -GroupScope Global -GroupCategory Security `
                             -Path $AdDomain.UsersContainer -PassThru
        Write-Ok "Created global security group '$TargetGroup'."
    }
    else {
        Write-Ok "Group '$TargetGroup' already exists - reusing it."
    }

    # ========================================================================
    # 2. ROOT FOLDER + NTFS PERMISSIONS
    # ========================================================================
    Write-Step "2/6  Root folder and NTFS permissions on '$LocalRootFolderPath'"
    if (-not (Test-Path -LiteralPath $LocalRootFolderPath)) {
        New-Item -Path $LocalRootFolderPath -ItemType Directory -Force | Out-Null
        Write-Ok "Created folder '$LocalRootFolderPath'."
    }
    else {
        Write-Ok "Folder '$LocalRootFolderPath' already exists."
    }

    $Acl = Get-Acl -LiteralPath $LocalRootFolderPath
    # Block inheritance and drop everything that was inherited
    $Acl.SetAccessRuleProtection($true, $false)

    # Strip leftover explicit rules so the final permission set is exactly ours (idempotent)
    foreach ($Rule in @($Acl.Access)) {
        [void]$Acl.RemoveAccessRule($Rule)
    }

    $Inherit = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit"
    $InheritOnly = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit"
    $NoPropagation = [System.Security.AccessControl.PropagationFlags]"None"
    $FullControl = [System.Security.AccessControl.FileSystemRights]"FullControl"
    $Allow = [System.Security.AccessControl.AccessControlType]"Allow"

    # Full control for administrators and SYSTEM, inherited by every user folder
    $AdminGroup = New-Object System.Security.Principal.NTAccount("$DomainNetBIOS\Domain Admins")
    $LocalAdmins = New-Object System.Security.Principal.SecurityIdentifier($SidLocalAdmins)
    $SystemAccount = New-Object System.Security.Principal.SecurityIdentifier($SidSystem)
    $CreatorOwner = New-Object System.Security.Principal.SecurityIdentifier($SidCreatorOwner)
    $GroupAccount = New-Object System.Security.Principal.NTAccount("$DomainNetBIOS\$TargetGroup")

    $Acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
        $AdminGroup, $FullControl, $Inherit, $NoPropagation, $Allow)))
    $Acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
        $LocalAdmins, $FullControl, $Inherit, $NoPropagation, $Allow)))
    $Acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
        $SystemAccount, $FullControl, $Inherit, $NoPropagation, $Allow)))

    # Target group: only enough rights to create and enter their OWN folder.
    # "This folder only" (no inheritance) means they cannot list or open other users' folders.
    $GroupRights = [System.Security.AccessControl.FileSystemRights]"CreateDirectories, Traverse, ReadAttributes, ReadPermissions"
    $Acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
        $GroupAccount, $GroupRights, [System.Security.AccessControl.InheritanceFlags]"None",
        $NoPropagation, $Allow)))

    # CREATOR OWNER: full control of the folder they create, inherited only
    $Acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
        $CreatorOwner, $FullControl, $InheritOnly,
        [System.Security.AccessControl.PropagationFlags]"InheritOnly", $Allow)))

    Set-Acl -LiteralPath $LocalRootFolderPath -AclObject $Acl
    Write-Ok "Inheritance disabled, old rules removed, 5 rules applied (admins/SYSTEM/creator-owner/group)."

    # ========================================================================
    # 3. SMB SHARE WITH ABE
    # ========================================================================
    Write-Step "3/6  SMB share '$ShareName' with Access-Based Enumeration"
    $AuthUsersName = Get-SidName -Sid $SidAuthenticatedUsers
    $Share = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
    if (-not $Share) {
        New-SmbShare -Name $ShareName -Path $LocalRootFolderPath -FullAccess $AuthUsersName `
                     -FolderEnumerationMode AccessBased | Out-Null
        Write-Ok "Created share '$ShareName' -> '$LocalRootFolderPath' with ABE on."
    }
    else {
        $SharePath = Get-SmbShare -Name $ShareName | Select-Object -ExpandProperty Path
        if ($SharePath -ne $LocalRootFolderPath) {
            Set-SmbShare -Name $ShareName -Path $LocalRootFolderPath -Force
            Write-Warn "Share pointed at '$SharePath' - repointed to '$LocalRootFolderPath'."
        }
        Set-SmbShare -Name $ShareName -FolderEnumerationMode AccessBased -Force
        Get-SmbShareAccess -Name $ShareName | Where-Object { $_.AccessControlType -eq 'Allow' } |
            ForEach-Object { Grant-SmbShareAccess -Name $ShareName -AccountName $_.AccountName -AccessRight Full -Force | Out-Null }
        Write-Ok "Share '$ShareName' already exists - ABE enforced, Full access kept."
    }

    # ========================================================================
    # 4. GPO: SELECT OR CREATE, FILTER, LINK
    # ========================================================================
    Write-Step "4/6  GPO '$GPOName'"
    $GPO = Get-GPO -Name $GPOName -ErrorAction SilentlyContinue
    $GpoNew = $false
    if (-not $GPO) {
        $GPO = New-GPO -Name $GPOName
        $GpoNew = $true
        Write-Ok "Created GPO '$GPOName' (GUID $($GPO.Id))."
    }
    else {
        Write-Ok "Reusing existing GPO '$GPOName' (GUID $($GPO.Id))."
    }

    # Security filtering: nobody gets "Apply" through Authenticated Users; only the target group does
    $Permissions = Get-GPPermission -Name $GPOName -All
    $AuthUsersPerm = $Permissions | Where-Object { $_.Trustee.Name -eq $AuthUsersName }
    if ($AuthUsersPerm -and $AuthUsersPerm.Permission -ne 'GpoRead') {
        Set-GPPermission -Name $GPOName -TargetName $AuthUsersName -TargetType Group `
                         -PermissionLevel GpoRead -Replace | Out-Null
        Write-Ok "Removed 'Apply' from $AuthUsersName (kept read access)."
    }
    $GroupPerm = $Permissions | Where-Object { $_.Trustee.Name -eq $TargetGroup }
    if (-not $GroupPerm -or $GroupPerm.Permission -ne 'GpoApply') {
        Set-GPPermission -Name $GPOName -TargetName $TargetGroup -TargetType Group `
                         -PermissionLevel GpoApply | Out-Null
        Write-Ok "Granted 'Apply Group Policy' to '$TargetGroup'."
    }
    else {
        Write-Ok "'$TargetGroup' already has 'Apply Group Policy'."
    }

    # OU link
    $Links = (Get-GPInheritance -Target $TargetOU -ErrorAction SilentlyContinue).GpoLinks
    $Linked = $Links | Where-Object { $_.DisplayName -eq $GPOName }
    if (-not $Linked) {
        New-GPLink -Name $GPOName -Target $TargetOU -LinkEnabled Yes | Out-Null
        Write-Ok "Linked GPO to '$TargetOU'."
    }
    else {
        Write-Ok "GPO already linked to '$TargetOU'."
    }

    $GPOGuid   = $GPO.Id.ToString("B").ToUpper()
    $SysvolGpo = "\\$DomainFQDN\SYSVOL\$DomainFQDN\Policies\$GPOGuid"
    $UserSide  = Join-Path $SysvolGpo "User"
    Write-Info "GPO folder: $SysvolGpo"

    # ========================================================================
    # 5. fdeploy.ini + Drives.xml IN SYSVOL
    # ========================================================================
    Write-Step "5/6  Writing Folder Redirection and Drive Map settings into SYSVOL"

    # ---- 5A: fdeploy.ini (UTF-16LE with BOM, as the protocol requires) ---------
    $FdeployDir  = Join-Path $UserSide "Documents & Settings"
    $FdeployPath = Join-Path $FdeployDir "fdeploy.ini"
    $Fdeploy1    = Join-Path $FdeployDir "fdeploy1.ini"
    if (-not (Test-Path -LiteralPath $FdeployDir)) {
        New-Item -Path $FdeployDir -ItemType Directory -Force | Out-Null
    }
    if (Test-Path -LiteralPath $Fdeploy1) {
        Write-Warn "fdeploy1.ini also exists in this GPO - clients read it BEFORE fdeploy.ini."
        Write-Warn "Check that GPO in the editor, or delete fdeploy1.ini, or the settings below are ignored."
    }

    $FolderTargets = [ordered]@{}
    $FolderTargets["Desktop"]   = @{ Guid = $DesktopGuid;   Path = $UncDesktopPath }
    $FolderTargets["Documents"] = @{ Guid = $DocumentsGuid; Path = $UncDocumentsPath }
    if ($RedirectDownloads) { $FolderTargets["Downloads"] = @{ Guid = $DownloadsGuid; Path = $UncDownloadsPath } }

    # [Version] header, then one section per redirected folder.
    # s-1-1-0 (Everyone) = the path applies to every user the policy applies to.
    # First-run tip: configure one folder by hand in a throw-away GPO, open the fdeploy.ini
    # it produces, and compare it with this file - that is the ground truth for your Windows build.
    $Sb = New-Object System.Text.StringBuilder
    [void]$Sb.AppendLine('[Version]')
    [void]$Sb.AppendLine('Signature="$CHICAGO$"')
    [void]$Sb.AppendLine('Advanced=0')
    foreach ($Key in @($FolderTargets.Keys)) {
        $Entry = $FolderTargets[$Key]
        [void]$Sb.AppendLine('')
        [void]$Sb.AppendLine("[$($Entry.Guid)]")
        [void]$Sb.AppendLine("s-1-1-0=$($Entry.Path)")
        [void]$Sb.AppendLine("Flags=$RedirectionFlags")
    }
    $FdeployContent = $Sb.ToString()

    $IniChanged = $true
    if (Test-Path -LiteralPath $FdeployPath) {
        if ([System.IO.File]::ReadAllText($FdeployPath).Trim() -eq $FdeployContent.Trim()) { $IniChanged = $false }
    }
    if ($IniChanged) {
        [System.IO.File]::WriteAllText($FdeployPath, $FdeployContent,
            (New-Object System.Text.UnicodeEncoding($false, $true)))
        Write-Ok "fdeploy.ini written (UTF-16LE): Desktop, Documents$(if ($RedirectDownloads) { ', Downloads' }) -> $UncUserBasePath"
    }
    else {
        Write-Ok "fdeploy.ini already had the same settings - left untouched."
    }

    # ---- 5B: Drives.xml (UTF-8 without BOM) ----------------------------------
    $DrivesDir  = Join-Path $UserSide "Preferences\Drives"
    $DrivesPath = Join-Path $DrivesDir "Drives.xml"
    if (-not (Test-Path -LiteralPath $DrivesDir)) {
        New-Item -Path $DrivesDir -ItemType Directory -Force | Out-Null
    }

    $Stamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    $DrivesXml = @"
<?xml version="1.0" encoding="utf-8"?>
<Drives clsid="{8FDDCC1A-0C3C-43cd-A6B4-71A6DF20DA8C}">
    <Drive clsid="{935D1B74-9CB8-4e3c-9914-7DD559B7A417}" name="${DriveLetter}:" status="${DriveLetter}:" image="2" changed="$Stamp" uid="$([guid]::NewGuid().ToString("B"))" bypassErrors="1">
        <Properties action="U" thisDrive="NOCHANGE" newDrive="${DriveLetter}:" allDrives="NOCHANGE" userName="" path="$UncUserBasePath" label="$DriveLabel" persistent="1" useLetter="1" letter="$DriveLetter"/>
    </Drive>
</Drives>
"@
    $XmlChanged = $true
    if (Test-Path -LiteralPath $DrivesPath) {
        $Existing = [System.IO.File]::ReadAllText($DrivesPath)
        if ($Existing -match ('path="' + [regex]::Escape($UncUserBasePath) + '"') -and
            $Existing -match ('letter="' + $DriveLetter + '"')) { $XmlChanged = $false }
    }
    if ($XmlChanged) {
        [System.IO.File]::WriteAllText($DrivesPath, $DrivesXml,
            (New-Object System.Text.UTF8Encoding($false)))
        Write-Ok "Drives.xml written: ${DriveLetter}: -> $UncUserBasePath (Update, persistent, label '$DriveLabel')."
    }
    else {
        Write-Ok "Drives.xml already maps ${DriveLetter}: to that path - left untouched."
    }

    # ---- 5C: register CSEs, bump version, sync GPT.INI -----------------------
    Write-Step "6/6  Registering client-side extensions and bumping the GPO version"
    $GpoDn  = "CN=$GPOGuid,CN=Policies,CN=System,$DomainDN"
    $GpoAdsi = [ADSI]"LDAP://$GpoDn"

    $CurrentCse = ""
    if ($GpoAdsi.Properties["gPCUserExtensionNames"].Value) {
        $CurrentCse = [string]$GpoAdsi.Properties["gPCUserExtensionNames"].Value
    }
    $CseChanged = $false
    foreach ($Cse in @($FolderRedirectionCSE, $DriveMapsCSE)) {
        # .Contains() compares literally - -like would treat [ ] as wildcards
        if (-not $CurrentCse.Contains($Cse)) {
            if ($CurrentCse.Length -gt 0 -and -not $CurrentCse.EndsWith("]")) { $CurrentCse += "]" }
            $CurrentCse += $Cse
            $CseChanged = $true
        }
    }

    if ($CseChanged) {
        $GpoAdsi.Properties["gPCUserExtensionNames"].Value = $CurrentCse
        Write-Ok "Added missing client-side extension(s) to gPCUserExtensionNames."
    }
    else {
        Write-Ok "Both client-side extensions were already registered."
    }

    if ($IniChanged -or $XmlChanged -or $CseChanged -or $GpoNew) {
        $OldVersion = [int]$GpoAdsi.Properties["versionNumber"].Value
        $NewVersion = $OldVersion + 65536   # user-side change = +65536
        $GpoAdsi.Properties["versionNumber"].Value = $NewVersion
        $GpoAdsi.CommitChanges()
        Write-Ok "GPO user version: $OldVersion -> $NewVersion (triggers re-apply on clients)."

        $GptIni = Join-Path $SysvolGpo "GPT.INI"
        if (Test-Path -LiteralPath $GptIni) {
            $Lines = Get-Content -LiteralPath $GptIni
            $NewLines = foreach ($Line in $Lines) {
                if ($Line -match '^\s*Version\s*=') { "Version=$NewVersion" } else { $Line }
            }
            [System.IO.File]::WriteAllText($GptIni, (($NewLines -join "`r`n") + "`r`n"),
                (New-Object System.Text.UTF8Encoding($false)))
            Write-Ok "GPT.INI synced to Version=$NewVersion."
        }
        else {
            Write-Warn "GPT.INI not found at $GptIni - clients may not notice the change."
        }
    }
    else {
        Write-Ok "Nothing changed this run - GPO version left as it was (fully idempotent)."
    }

    # ---- 5D: wait for SYSVOL replication ------------------------------------
    if ($SysvolWaitSeconds -gt 0) {
        $DcHosts = @(Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName)
        if ($DcHosts.Count -gt 1) {
            Write-Info "Waiting up to $SysvolWaitSeconds s for SYSVOL to reach $($DcHosts.Count) DCs..."
            $FilesToCheck = @("User\Documents & Settings\fdeploy.ini", "User\Preferences\Drives\Drives.xml")
            $Deadline = $StartTime.AddSeconds($SysvolWaitSeconds)
            $Pending = $DcHosts
            while ($Pending.Count -gt 0 -and (Get-Date) -lt $Deadline) {
                $Pending = @($Pending | Where-Object {
                    $Dc = $_
                    @($FilesToCheck | Where-Object {
                        -not (Test-Path -LiteralPath "\\$Dc\SYSVOL\$DomainFQDN\Policies\$GPOGuid\$_")
                    }).Count -gt 0
                })
                if ($Pending.Count -gt 0) { Start-Sleep -Seconds 5 }
            }
            if ($Pending.Count -eq 0) {
                Write-Ok "SYSVOL settings copied to all domain controllers."
            }
            else {
                Write-Warn "Still copying to: $($Pending -join ', '). Check with repadmin /showrepl later."
            }
        }
    }

    # ---- Done ---------------------------------------------------------------
    Write-Host "`n=== Deployment finished in $([int]((Get-Date) - $StartTime).TotalSeconds)s ===" -ForegroundColor Green
    Write-Host "GPO            : $GPOName ($GPOGuid)"
    Write-Host "Applies to     : members of '$TargetGroup' (users must be inside $TargetOU for the link)"
    Write-Host "Redirected to  : $UncUserBasePath"
    Write-Host "Drive mapped   : ${DriveLetter}: ($DriveLabel)"
    Write-Host "`nHow to verify:" -ForegroundColor Cyan
    Write-Info "1. Standard user: add a test user to '$TargetGroup', sign in on a client, run 'gpupdate /force', sign out and back in."
    Write-Info "   Desktop / Documents should show '$UncUserBasePath' in their folder properties; the drive should appear."
    Write-Info "2. That user must NOT be able to open another user's folder under $UncRootPath (ABE + NTFS block this)."
    Write-Info "3. Domain Admin must be able to browse every folder under $UncRootPath (exclusive rights are off)."
    Write-Info "4. Open the GPO in the group policy editor and confirm the three folders show a UNC path there too."
    Write-Info "5. Client troubleshooting: 'gpresult /h report.html' then check Folder Redirection and Drive Maps sections."
}
catch {
    Write-Host "`n=== Deployment FAILED ===" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "At line: $($_.InvocationInfo.ScriptLineNumber)  ->  $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkYellow
    exit 1
}
