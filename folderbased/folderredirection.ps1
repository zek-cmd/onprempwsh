<#
.SYNOPSIS
    Builds isolated per-user home folders on a file server: root folder with strict NTFS,
    hidden SMB share with Access-Based Enumeration (ABE), a Group Policy Object with
    Folder Redirection (Desktop / Documents / Downloads) and a mapped drive, all built
    on dynamic %USERNAME% paths.

.DESCRIPTION
    Steps:
      1. Ask WHERE the redirection folder must live - you pick it in a standard Windows
         folder dialog (the same style as Open / Save As) - create it only after you say
         yes, and lock its NTFS permissions down.
      2. Create (or reuse) the hidden SMB share, with ABE on.
      3. Ask WHICH EXISTING GPO to use - this script never creates a GPO - and link it
         to the target OU.
      4. Write fdeploy1.ini (Folder Redirection) and Drives.xml (GPP Drive Maps) into
         SYSVOL, register both client-side extensions in AD, bump the GPO user version
         and sync GPT.INI.
      5. Wait for SYSVOL to reach the other domain controllers and print how to verify.

    No security group is created. The policy is scoped by the OU link alone, and every
    user the GPO applies to is redirected into their own %USERNAME% folder.

    Every step is idempotent: run it twice and the second run changes nothing.

.NOTES
    Requires: Domain Admin rights, RSAT (ActiveDirectory + GroupPolicy modules).
    Run it as Administrator on the file server, or on a management PC that can reach
    the server's local disk path and SYSVOL over the network.
    Run it in Windows PowerShell 5.1 (powershell.exe). If you start it from
    PowerShell 7 (pwsh.exe) the script hands the run over to 5.1 by itself.
#>

#Requires -RunAsAdministrator

# ============================================================================
# RUNTIME GUARD - START IN THE RIGHT POWERSHELL
# ============================================================================
# The Group Policy cmdlets (Get-GPO, Set-GPPermission, New-GPLink ...) live only in
# Windows PowerShell 5.1. PowerShell 7 cannot load that module, so before the old
# "#Requires -Modules GroupPolicy" line it would stop with a missing-module error.
# Started from PowerShell 7, the script now re-runs itself in 5.1 instead.

if ($PSVersionTable.PSEdition -eq 'Core') {
    $WinPsExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $WinPsExe)) {
        throw "This script needs the Group Policy cmdlets, which only exist in Windows PowerShell 5.1. powershell.exe was not found at '$WinPsExe'."
    }
    Write-Host "PowerShell 7 cannot load the Group Policy module." -ForegroundColor Yellow
    Write-Host "Re-running this script in Windows PowerShell 5.1 (powershell.exe)..." -ForegroundColor Yellow
    & $WinPsExe -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath
    exit $LASTEXITCODE
}

foreach ($RequiredModule in @('ActiveDirectory', 'GroupPolicy')) {
    if (-not (Get-Module -ListAvailable -Name $RequiredModule)) {
        $HowToInstall = if ((Get-CimInstance -ClassName Win32_OperatingSystem).ProductType -eq 1) {
            "Add-WindowsCapability -Online -Name Rsat.$RequiredModule.Tools~~~~0.0.1.0"
        }
        else {
            "Install-WindowsFeature RSAT-AD-PowerShell, RSAT-GPMC"
        }
        throw "The '$RequiredModule' module is not installed on '$($env:COMPUTERNAME)'. This script needs RSAT. Install it with:  $HowToInstall   then run the script again."
    }
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================================
# 1. MANUAL CONFIGURATION - EDIT THESE TO MATCH YOUR ENVIRONMENT
# ============================================================================
$FileServerName        = ""                                      # Server holding the share. "" = the PC the script runs on
$LocalRootFolderPath   = ""                                      # Asked when the script starts. "" = ask (suggestion below)
$DefaultRootFolderPath = "D:\UserRedirection"                    # Suggested answer for the question above
$ShareName             = "UserRedirection$"                     # SMB share name ($ = hidden share)

# Subfolder names created inside each user's %USERNAME% directory
$DesktopFolderName     = "Desktop"
$DocumentsFolderName   = "Documents"
$DownloadsFolderName   = "Downloads"

# Active Directory & Group Policy settings
$DomainFQDN            = ""                                      # e.g. contoso.com - leave "" to auto-detect
$DomainNetBIOS         = ""                                      # e.g. CONTOSO - leave "" to auto-detect
$GPOName               = ""                                      # "" = the script lists existing GPOs and you pick one. It never creates one.
$TargetOU              = ""                                      # e.g. OU=Users,DC=contoso,DC=com - "" = ask at run time (domain root or any OU)

# Drive mapping settings
$DriveLetter           = "Z"                                    # Drive letter, colon optional
$DriveLabel            = "Personal Network Drive"               # Label shown in File Explorer

# Folder redirection behaviour, written in the Flags= value of fdeploy1.ini.
# Plain hexadecimal digits, exactly as the Group Policy editor writes them (no "0x").
# 1001 = 0x1001: move the existing contents into the per-user folder (0x1) plus the
#        bit Windows' own editor always adds (0x1000). Domain Admins / SYSTEM keep
#        access to the folders - this is the editor's "Grant the user exclusive
#        rights" box UNTICKED.
# Never combine "Redirection not specified" (0x4) with a FullPath - the editor rejects it.
$RedirectionFlags      = "1001"

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

# AppData is left unredirected, but the editor still writes its line with Flags=4
$AppDataGuid           = "{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}"

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

function Read-Answer {
    param([string]$Prompt, [string]$Default = "")
    while ($true) {
        if ([string]::IsNullOrWhiteSpace($Default)) { $Answer = Read-Host $Prompt }
        else { $Answer = Read-Host "$Prompt [$Default]" }
        if (-not [string]::IsNullOrWhiteSpace($Answer)) { return $Answer.Trim() }
        if (-not [string]::IsNullOrWhiteSpace($Default)) { return $Default }
        Write-Warn "Please type an answer."
    }
}

function Read-YesNo {
    param([string]$Prompt)
    while ($true) {
        $Answer = Read-Host "$Prompt [y/N]"
        if ([string]::IsNullOrWhiteSpace($Answer)) { return $false }
        if ($Answer.Trim() -match '^(y|yes)$') { return $true }
        if ($Answer.Trim() -match '^(n|no)$') { return $false }
        Write-Warn "Answer y (yes) or n (no)."
    }
}

function Show-FolderPicker {
    # Windows folder dialog, so the redirection folder is chosen with the mouse instead
    # of typed into PowerShell. New Folder works inside it.
    #   1) the modern "Select Folder" window (the same family as Open / Save As),
    #   2) the classic shell browse window if that one cannot load,
    #   3) $null when it was closed or cancelled, so the caller can offer a typed path.
    param(
        [string]$Description = 'Select the folder that will hold the redirected user folders',
        [string]$InitialPath = ''
    )

    # --- 1) Modern folder dialog -------------------------------------------------
    $DialogProblem = ''
    try {
        if ($null -eq ('FolderPickerNative' -as [type])) {
            $PickerSource = @'
using System;
using System.Runtime.InteropServices;

public class FolderPickerNative
{
    [ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IFileDialog
    {
        [PreserveSig] int Show(IntPtr parent);
        void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise(IntPtr pfde, out uint pdwCookie);
        void Unadvise(uint dwCookie);
        void SetOptions(uint fos);
        void GetOptions(out uint pfos);
        void SetDefaultFolder(IShellItem psi);
        void SetFolder(IShellItem psi);
        void GetFolder(out IShellItem ppsi);
        void GetCurrentSelection(out IShellItem ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult(out IShellItem ppsi);
        void AddPlace(IShellItem psi, int fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close(int hr);
        void SetClientGuid(ref Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr pFilter);
    }

    [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellItem
    {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHCreateItemFromParsingName(string pszPath, IntPtr pbc, ref Guid riid, out IShellItem ppv);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    public static string Pick(string title, string okLabel, string initialPath)
    {
        object rawDialog = Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")));
        IFileDialog dialog = (IFileDialog)rawDialog;

        try
        {
            // folders only, real file-system folders only, and the path must exist
            dialog.SetOptions(0x20u | 0x40u | 0x800u);
            dialog.SetTitle(title);
            dialog.SetOkButtonLabel(okLabel);

            if (!string.IsNullOrEmpty(initialPath))
            {
                Guid shellItemId = typeof(IShellItem).GUID;
                IShellItem startFolder;
                if (SHCreateItemFromParsingName(initialPath, IntPtr.Zero, ref shellItemId, out startFolder) == 0 && startFolder != null)
                {
                    dialog.SetFolder(startFolder);
                    Marshal.ReleaseComObject(startFolder);
                }
            }

            // 0 = the user picked a folder; anything else = closed or cancelled
            if (dialog.Show(GetForegroundWindow()) != 0) { return string.Empty; }

            IShellItem result;
            dialog.GetResult(out result);
            if (result == null) { return string.Empty; }

            string path;
            result.GetDisplayName(0x80058000u, out path);
            Marshal.ReleaseComObject(result);
            return path == null ? string.Empty : path;
        }
        finally
        {
            if (Marshal.IsComObject(rawDialog)) { Marshal.ReleaseComObject(rawDialog); }
        }
    }
}
'@
            Add-Type -TypeDefinition $PickerSource -Language CSharp -ErrorAction Stop
        }

        $ModernPath = [FolderPickerNative]::Pick($Description, 'Select Folder', $InitialPath)
        if (-not [string]::IsNullOrWhiteSpace($ModernPath)) { return $ModernPath.Trim() }
        return $null
    }
    catch {
        # the modern dialog could not load - try the classic one below
        $DialogProblem = $_.Exception.Message
    }

    # --- 2) Classic shell browse dialog -----------------------------------------
    try {
        $Shell = New-Object -ComObject Shell.Application
        # 0x40 = new dialog style (gives the New Folder button), 0x10 = show an edit box
        $BrowseFlags = 0x40 -bor 0x10
        $Chosen = $Shell.BrowseForFolder(0, $Description, $BrowseFlags, $InitialPath)
        if ($null -ne $Chosen) {
            $ChosenPath = [string]$Chosen.Self.Path
            if (-not [string]::IsNullOrWhiteSpace($ChosenPath)) { return $ChosenPath.Trim() }
        }
    }
    catch {
        # fall through and let the caller offer a typed path
        $DialogProblem = "$DialogProblem $($_.Exception.Message)".Trim()
    }

    if (-not [string]::IsNullOrWhiteSpace($DialogProblem)) {
        Write-Warn "The folder dialog could not open: $DialogProblem"
    }
    return $null
}

function Select-ExistingGpo {
    $AllGpo = @(Get-GPO -All | Sort-Object DisplayName)
    if ($AllGpo.Count -eq 0) {
        throw "No GPO exists in this domain yet. Create the GPO you want in Group Policy Management (or the Group Policy editor), then run this script again."
    }

    Write-Host "`nExisting GPOs - pick one, this script will not create a new one:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $AllGpo.Count; $i++) {
        Write-Host ("  [{0}] {1}" -f ($i + 1), $AllGpo[$i].DisplayName)
    }

    while ($true) {
        $Pick = Read-Answer -Prompt "Number of the GPO to use, or its exact name"
        if ($Pick -match '^\d+$') {
            $Index = [int]$Pick
            if ($Index -ge 1 -and $Index -le $AllGpo.Count) { return $AllGpo[$Index - 1].DisplayName }
            Write-Warn "There is no GPO number $Index in that list."
            continue
        }
        $Match = $AllGpo | Where-Object { $_.DisplayName -eq $Pick }
        if ($Match) { return $Match.DisplayName }
        Write-Warn "No GPO is named '$Pick'. Type the number from the list, or the exact name."
    }
}

function Test-GpoLinkTarget {
    # A GPO can only be linked to the domain root or to an OU. A plain container such
    # as CN=Users (AD's default Users container) is refused by New-GPLink with the
    # unhelpful message "There is no such object on the server".
    param([string]$DistinguishedName, [string]$DomainDN)
    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return $false }
    $DN = $DistinguishedName.Trim()
    if ($DN -eq $DomainDN) { return $true }
    $Ou = Get-ADOrganizationalUnit -Identity $DN -ErrorAction SilentlyContinue
    return [bool]$Ou
}

function Select-GpoLinkTarget {
    param($Domain)

    Write-Host "`nWhere must this GPO apply? A GPO links to the whole domain, or to one OU:" -ForegroundColor Cyan
    Write-Host ("  [1] Whole domain '{0}' - every user in the domain" -f $Domain.DNSRoot)
    $AllOu = @(Get-ADOrganizationalUnit -Filter * -ErrorAction SilentlyContinue | Sort-Object Name)
    for ($i = 0; $i -lt $AllOu.Count; $i++) {
        Write-Host ("  [{0}] {1}" -f ($i + 2), $AllOu[$i].Name)
        Write-Host ("       {0}" -f $AllOu[$i].DistinguishedName) -ForegroundColor DarkGray
    }

    while ($true) {
        $Pick = Read-Answer -Prompt 'Number of the target from the list, or its exact distinguished name' -Default '1'
        if ($Pick -match '^\d+$') {
            $Index = [int]$Pick
            if ($Index -eq 1) { return [string]$Domain.DistinguishedName }
            if ($Index -ge 2 -and $Index -le ($AllOu.Count + 1)) { return [string]$AllOu[$Index - 2].DistinguishedName }
            Write-Warn "There is no target number $Index in that list."
            continue
        }
        if (Test-GpoLinkTarget -DistinguishedName $Pick -DomainDN $Domain.DistinguishedName) { return $Pick.Trim() }
        Write-Warn "'$Pick' is neither the domain root nor an existing OU. Pick a number from the list, or paste an OU distinguished name."
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
    if ([string]::IsNullOrWhiteSpace($TargetOU)) {
        # The old default was the Users container, which cannot hold a GPO link at all.
        $TargetOU = Select-GpoLinkTarget -Domain $AdDomain
    }
    $TargetOU = $TargetOU.Trim()
    if (-not (Test-GpoLinkTarget -DistinguishedName $TargetOU -DomainDN $DomainDN)) {
        Write-Warn "'$TargetOU' is not the domain root and not an OU, so Windows cannot link a GPO there."
        Write-Warn 'A plain container such as CN=Users is the usual mistake - that is what made the earlier link fail.'
        $TargetOU = Select-GpoLinkTarget -Domain $AdDomain
    }

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
    Write-Info "Detected: GPO link target '$TargetOU'"
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

    $AuthUsersName = Get-SidName -Sid $SidAuthenticatedUsers

    # ========================================================================
    # ASK WHERE THE REDIRECTION FOLDER GOES - nothing is created on its own
    # ========================================================================
    if (-not [System.Environment]::UserInteractive) {
        throw "This script asks questions, so it must run in an interactive PowerShell window. Unattended runs are not supported."
    }

    if ([string]::IsNullOrWhiteSpace($LocalRootFolderPath)) {
        $InitialPickPath = Split-Path -Parent $DefaultRootFolderPath
        if ([string]::IsNullOrWhiteSpace($InitialPickPath) -or -not (Test-Path -LiteralPath $InitialPickPath)) {
            $InitialPickPath = $env:SystemDrive
        }

        Write-Info 'Opening the folder dialog - pick the folder that will hold the user folders (New Folder works there too).'
        $Picked = Show-FolderPicker -Description "Select the folder that will hold the redirected user folders (use 'New Folder' to create it)" -InitialPath $InitialPickPath

        if ([string]::IsNullOrWhiteSpace($Picked)) {
            if (-not (Read-YesNo -Prompt 'No folder was chosen in the dialog. Type the path by hand instead?')) {
                throw 'Stopped: no redirection folder was chosen, so nothing was created or changed.'
            }
            $Picked = Read-Answer -Prompt 'Where must the redirection folder live? Type a drive path (D:\UserRedirection) or a UNC path' -Default $DefaultRootFolderPath
        }
        $LocalRootFolderPath = $Picked
    }
    $LocalRootFolderPath = $LocalRootFolderPath.Trim().Trim('"')
    if ($LocalRootFolderPath -notmatch '^([A-Za-z]:\\|\\\\)') {
        throw "'$LocalRootFolderPath' is not a folder path. Use a drive path such as D:\UserRedirection, or a UNC path such as \\\\FILESERVER\\UserRedirection$."
    }

    Write-Info "Redirection folder chosen: $LocalRootFolderPath"

    # A bare drive root (D:\) or a bare share root (\\SERVER\Share) would drop every user
    # folder straight onto that root. Offer the usual folder name instead, and let you keep
    # the root if that is really what you meant.
    if ($LocalRootFolderPath -match '^[A-Za-z]:\\?$' -or $LocalRootFolderPath -match '^\\\\[^\\]+\\[^\\]+\\?$') {
        $RootLeafName = Split-Path -Leaf $DefaultRootFolderPath
        $RootSuggestion = "$($LocalRootFolderPath.TrimEnd('\'))\$RootLeafName"
        Write-Warn "You picked the root '$LocalRootFolderPath'. Every user folder would sit directly there."
        if (Read-YesNo -Prompt "Use '$RootSuggestion' instead?") { $LocalRootFolderPath = $RootSuggestion }
        Write-Info "Redirection folder chosen: $LocalRootFolderPath"
    }
    if (-not (Test-Path -LiteralPath $LocalRootFolderPath)) {
        Write-Warn "That folder does not exist yet: $LocalRootFolderPath"
        if (-not (Read-YesNo -Prompt "Create it now?")) {
            throw "Stopped: the folder was not created and no share, GPO or setting was touched."
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

    Write-Host ""

    # ========================================================================
    # 1. ROOT FOLDER + NTFS PERMISSIONS
    # ========================================================================
    Write-Step "1/5  Root folder and NTFS permissions on '$LocalRootFolderPath'"
    $FolderWasNew = $false
    if (-not (Test-Path -LiteralPath $LocalRootFolderPath)) {
        New-Item -Path $LocalRootFolderPath -ItemType Directory -Force | Out-Null
        $FolderWasNew = $true
        Write-Ok "Created folder '$LocalRootFolderPath'."
    }
    else {
        Write-Ok "Folder '$LocalRootFolderPath' already exists."
        Write-Warn "The next step replaces its NTFS permissions: inheritance off, old rules removed."
        if (-not (Read-YesNo -Prompt "Replace the permissions on this existing folder?")) {
            throw "Stopped: permissions were left exactly as they were, and nothing else was touched."
        }
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
    $AuthUsers = New-Object System.Security.Principal.SecurityIdentifier($SidAuthenticatedUsers)

    $Acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
        $AdminGroup, $FullControl, $Inherit, $NoPropagation, $Allow)))
    $Acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
        $LocalAdmins, $FullControl, $Inherit, $NoPropagation, $Allow)))
    $Acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
        $SystemAccount, $FullControl, $Inherit, $NoPropagation, $Allow)))

    # Every signed-in user: only enough rights to create and enter their OWN folder.
    # "This folder only" (no inheritance) means they cannot list or open other users' folders.
    $AuthUsersRights = [System.Security.AccessControl.FileSystemRights]"CreateDirectories, Traverse, ReadAttributes, ReadPermissions"
    $Acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
        $AuthUsers, $AuthUsersRights, [System.Security.AccessControl.InheritanceFlags]"None",
        $NoPropagation, $Allow)))

    # CREATOR OWNER: full control of the folder they create, inherited only
    $Acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
        $CreatorOwner, $FullControl, $InheritOnly,
        [System.Security.AccessControl.PropagationFlags]"InheritOnly", $Allow)))

    Set-Acl -LiteralPath $LocalRootFolderPath -AclObject $Acl
    Write-Ok "Inheritance disabled, old rules removed, 5 rules applied (admins/SYSTEM/creator-owner/authenticated users)."

    # ========================================================================
    # 3. SMB SHARE WITH ABE
    # ========================================================================
    Write-Step "2/5  SMB share '$ShareName' with Access-Based Enumeration"
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
    # 3. PICK AN EXISTING GPO AND LINK IT TO THE OU
    # ========================================================================
    Write-Step "3/5  Existing GPO"
    if ([string]::IsNullOrWhiteSpace($GPOName)) {
        $GPOName = Select-ExistingGpo
    }

    $GPO = Get-GPO -Name $GPOName -ErrorAction SilentlyContinue
    if (-not $GPO) {
        throw "There is no GPO named '$GPOName' in '$DomainFQDN'. This script never creates one - create it in Group Policy Management first, then run this again."
    }
    $GpoNew = $false
    Write-Ok "Using existing GPO '$GPOName' (GUID $($GPO.Id))."

    # Security filtering: the OU link is what scopes the policy, so Authenticated Users keeps "Apply".
    # (An earlier run of the old group-based version may have set it to read-only; this puts it back.)
    $Permissions = Get-GPPermission -Name $GPOName -All
    $AuthUsersPerm = $Permissions | Where-Object { $_.Trustee.Name -eq $AuthUsersName }
    if (-not $AuthUsersPerm -or $AuthUsersPerm.Permission -ne 'GpoApply') {
        Set-GPPermission -Name $GPOName -TargetName $AuthUsersName -TargetType Group `
                         -PermissionLevel GpoApply -Replace | Out-Null
        Write-Ok "Restored 'Apply Group Policy' for $AuthUsersName."
    }
    else {
        Write-Ok "$AuthUsersName may already apply this GPO."
    }

    # OU link - read the existing links defensively, so a failed read never aborts the run
    $Linked = $false
    try {
        $Inheritance = Get-GPInheritance -Target $TargetOU -ErrorAction Stop
        if ($Inheritance -and $Inheritance.PSObject.Properties['GpoLinks']) {
            $Linked = @($Inheritance.GpoLinks | Where-Object { $_.DisplayName -eq $GPOName }).Count -gt 0
        }
    }
    catch {
        Write-Warn "Could not read the existing links on '$TargetOU' ($($_.Exception.Message)). The link step tries anyway."
    }

    if (-not $Linked) {
        Write-Host "`nThe GPO '$GPOName' will now be linked to:" -ForegroundColor Cyan
        Write-Host "  $TargetOU"
        Write-Host "  (this makes the policy apply to every user in that scope)" -ForegroundColor DarkGray
        if (-not (Read-YesNo -Prompt 'Link it now?')) {
            throw "Stopped: the GPO was not linked, so no policy setting was changed."
        }
        try {
            New-GPLink -Name $GPOName -Target $TargetOU -LinkEnabled Yes -ErrorAction Stop | Out-Null
            Write-Ok "Linked GPO to '$TargetOU'."
        }
        catch {
            if ($_.Exception.Message -match 'already (exists|linked|in place)') {
                Write-Ok "GPO already linked to '$TargetOU'."
            }
            else { throw }
        }
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
    Write-Step "4/5  Writing Folder Redirection and Drive Map settings into SYSVOL"

    # ---- 5A: fdeploy1.ini (Format Version One, UTF-16LE with BOM) -------------
    $FdeployDir   = Join-Path $UserSide "Documents & Settings"
    if (-not (Test-Path -LiteralPath $FdeployDir)) {
        New-Item -Path $FdeployDir -ItemType Directory -Force | Out-Null
    }
    $Fdeploy1Path = Join-Path $FdeployDir "fdeploy1.ini"
    $FdeployZero   = Join-Path $FdeployDir "fdeploy.ini"

    $FolderTargets = [ordered]@{}
    $FolderTargets["Desktop"]   = @{ Guid = $DesktopGuid;   Path = $UncDesktopPath }
    $FolderTargets["Documents"] = @{ Guid = $DocumentsGuid; Path = $UncDocumentsPath }
    if ($RedirectDownloads) { $FolderTargets["Downloads"] = @{ Guid = $DownloadsGuid; Path = $UncDownloadsPath } }

    # fdeploy1.ini = Format Version One, and clients read it BEFORE the old fdeploy.ini.
    # The shape below is copied from a GPO that the Group Policy editor produced on this
    # domain itself: [version] version=100, one [Folder_Redirection] line per folder,
    # then one [GUID_s-1-1-0] section holding Flags= and FullPath=.
    # Never write [GUID] sections into fdeploy.ini - that is Format Version Zero, which
    # has no such keys, and the editor then reports "FRSettingRead failed with -2147020590".
    $IniSid = "s-1-1-0"     # every user the policy applies to, lower case as Windows writes it
    $Sb = New-Object System.Text.StringBuilder
    [void]$Sb.AppendLine('[version]')
    [void]$Sb.AppendLine('version=100')
    [void]$Sb.AppendLine('[Folder_Redirection]')
    [void]$Sb.AppendLine("$AppDataGuid=$IniSid;")
    foreach ($Key in @($FolderTargets.Keys)) {
        [void]$Sb.AppendLine("$($FolderTargets[$Key].Guid)=$IniSid;")
    }
    [void]$Sb.AppendLine('')
    [void]$Sb.AppendLine("[$AppDataGuid`_$IniSid]")
    [void]$Sb.AppendLine('Flags=4')
    foreach ($Key in @($FolderTargets.Keys)) {
        $Entry = $FolderTargets[$Key]
        [void]$Sb.AppendLine('')
        [void]$Sb.AppendLine("[$($Entry.Guid)_$IniSid]")
        [void]$Sb.AppendLine("Flags=$RedirectionFlags")
        [void]$Sb.AppendLine("FullPath=$($Entry.Path)")
    }
    $Fdeploy1Content = $Sb.ToString().TrimEnd() + "`r`n"

    $IniChanged = $true
    if (Test-Path -LiteralPath $Fdeploy1Path) {
        $OldIni = [System.IO.File]::ReadAllText($Fdeploy1Path)
        if ($OldIni.Trim() -eq $Fdeploy1Content.Trim()) { $IniChanged = $false }
        elseif ($OldIni.Trim().Length -gt 0) {
            [System.IO.File]::WriteAllBytes("$Fdeploy1Path.bak", [System.IO.File]::ReadAllBytes($Fdeploy1Path))
            Write-Warn "This GPO already had a different fdeploy1.ini - old copy saved as fdeploy1.ini.bak."
        }
    }
    if ($IniChanged) {
        [System.IO.File]::WriteAllText($Fdeploy1Path, $Fdeploy1Content,
            (New-Object System.Text.UnicodeEncoding($false, $true)))
        Write-Ok "fdeploy1.ini written (Format Version One, UTF-16LE): Desktop, Documents$(if ($RedirectDownloads) { ', Downloads' }) -> $UncUserBasePath"
    }
    else {
        Write-Ok "fdeploy1.ini already had the same settings - left untouched."
    }

    # The legacy fdeploy.ini must stay a whitespace-only placeholder, the way Windows
    # keeps it next to fdeploy1.ini. Old-format content there breaks the policy editor.
    if (Test-Path -LiteralPath $FdeployZero) {
        $ZeroText = ""
        try { $ZeroText = [System.IO.File]::ReadAllText($FdeployZero) } catch { $ZeroText = "" }
        if ($ZeroText.Trim().Length -gt 0) {
            [System.IO.File]::WriteAllText($FdeployZero, "`r`n     `r`n",
                (New-Object System.Text.UnicodeEncoding($false, $true)))
            Write-Ok "Old-format content cleared out of fdeploy.ini - that is what the policy editor choked on."
        }
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
    Write-Step "5/5  Registering client-side extensions and bumping the GPO version"
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
            $FilesToCheck = @("User\Documents & Settings\fdeploy1.ini", "User\Documents & Settings\fdeploy.ini", "User\Preferences\Drives\Drives.xml")
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
    Write-Host "Applies to     : every user in '$TargetOU' (the GPO link is what scopes the policy, no security group)"
    Write-Host "Redirected to  : $UncUserBasePath"
    Write-Host "Drive mapped   : ${DriveLetter}: ($DriveLabel)"
    Write-Host "`nHow to verify:" -ForegroundColor Cyan
    Write-Info "1. Sign in on a client as a test user from '$TargetOU', run 'gpupdate /force', sign out and back in."
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
