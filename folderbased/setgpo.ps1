<#
.SYNOPSIS
    Hardens workstations through an EXISTING Group Policy Object: locks regedit, wallpaper,
    lock screen image and the taskbar, turns off the Microsoft Store and Offline Files,
    stops Windows caching domain logons, turns on verbose logon messages, deletes stale
    user profiles after 30 days, enables WinRM (plus compatibility mode) and fills the
    GPO with a baseline set of inbound / outbound firewall rules, including ICMP so
    domain-wide ping works.

.DESCRIPTION
    Steps:
      1. Pick an existing GPO (this script never creates one) and link it to the whole
         domain or to one OU, the same way folderredirection.ps1 asks.
      2. Show every setting the script is about to write, then wait for your yes.
      3. Write the User Configuration and Computer Configuration registry policies into
         the GPO's registry.pol through the Group Policy cmdlets - each setting is read
         first, so a second run changes nothing.
      4. Merge the two Security Policy entries into GptTmpl.inf (CachedLogonsCount and the
         WinRM service start type), keeping every other line of that file untouched.
      5. Create the firewall rules inside the GPO itself, so clients receive them.
      6. Bump the GPO version only for the file-based changes, sync GPT.INI, wait for SYSVOL.

.NOTES
    Requires: Domain Admin rights, RSAT (ActiveDirectory + GroupPolicy modules), and
    Windows 8 / Server 2012 or newer on this PC (the firewall cmdlets that can edit a GPO
    need it).
    Run it in Windows PowerShell 5.1. Started from PowerShell 7 it hands the run to 5.1.
#>

#Requires -RunAsAdministrator

# ============================================================================
# RUNTIME GUARD - START IN THE RIGHT POWERSHELL
# ============================================================================
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

# Which GPO and where it applies. "" = the script lists existing GPOs and you pick one.
# It NEVER creates a GPO. "" for the OU = the script lists the domain and every OU.
$GPOName               = ""
$TargetOU              = ""
$DomainNetBIOS         = ""            # "" = auto-detect from Active Directory

# Each switch below turns one policy on ($true) or leaves it alone ($false).
$DisableRegedit        = $true         # Regedit.exe refuses to start
$DisableWallpaperChange = $true        # Desktop background locked
$DisableLockscreenChange = $true       # Lock screen / logon image locked
$DisableTaskbarChange  = $true         # Taskbar settings locked
$DisableMicrosoftStore = $true         # Store app and Store-sourced apps blocked
$DisableOfflineFiles   = $true         # Offline Files turned off
$EnableVerboseStatus   = $true         # "Display highly detailed status messages"
$DisableCredentialCache = $true        # No cached domain logons (CachedLogonsCount = 0)
$EnableProfileCleanup  = $true         # Delete stale user profiles on restart
$ProfileCleanupDays    = 30            # How old a profile must be before deletion

# WinRM + firewall
$EnableWinRM           = $true         # WinRM service auto-start + listens on the network
$EnableWinRMCompat     = $true         # WinRM 1.1 compatibility rule (TCP 80 inbound)
$EnableWinRMBasicAuth  = $false        # Leave $false unless a legacy tool needs it
$EnableFirewallRules   = $true         # Write the rule list below into the GPO
$EnableRdpInbound      = $false        # Remote Desktop rule (TCP 3389 inbound)

# Waits for SYSVOL to reach the other domain controllers (0 = do not wait)
$SysvolWaitSeconds     = 90

# ============================================================================
# 2. POLICY SETTINGS - plain lists you can edit
# ============================================================================
# Key      : registry policy key. HKCU goes to User Configuration, HKLM to Computer
#            Configuration. ValueName / Type / Value speak for themselves.
# Scope is worked out from the key prefix, so there is no extra field to get wrong.
$UserPolicySettings = @(
    @{
        Setting  = 'Disable registry editing tools (regedit)'
        Key      = 'HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System'
        ValueName = 'DisableRegistryTools'
        Type     = 'DWord'
        Value    = 2
        Enabled  = $DisableRegedit
    }
    @{
        Setting  = 'Prevent changing desktop background'
        Key      = 'HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop'
        ValueName = 'NoChangingWallPaper'
        Type     = 'DWord'
        Value    = 1
        Enabled  = $DisableWallpaperChange
    }
    @{
        Setting  = 'Lock taskbar settings'
        Key      = 'HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer'
        ValueName = 'TaskbarLockAll'
        Type     = 'DWord'
        Value    = 1
        Enabled  = $DisableTaskbarChange
    }
)

$ComputerPolicySettings = @(
    @{
        Setting  = 'Prevent changing lock screen and logon image'
        Key      = 'HKLM\Software\Policies\Microsoft\Windows\Personalization'
        ValueName = 'NoChangingLockScreen'
        Type     = 'DWord'
        Value    = 1
        Enabled  = $DisableLockscreenChange
    }
    @{
        Setting  = 'Turn off the Microsoft Store application'
        Key      = 'HKLM\Software\Policies\Microsoft\WindowsStore'
        ValueName = 'RemoveWindowsStore'
        Type     = 'DWord'
        Value    = 1
        Enabled  = $DisableMicrosoftStore
    }
    @{
        Setting  = 'Allow or Disallow use of the Offline Files feature (Disabled)'
        Key      = 'HKLM\Software\Policies\Microsoft\Windows\NetCache'
        ValueName = 'Enabled'
        Type     = 'DWord'
        Value    = 0
        Enabled  = $DisableOfflineFiles
    }
    @{
        Setting  = 'Display highly detailed status messages'
        Key      = 'HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System'
        ValueName = 'VerboseStatus'
        Type     = 'DWord'
        Value    = 1
        Enabled  = $EnableVerboseStatus
    }
    @{
        Setting  = 'Delete user profiles older than N days on restart'
        Key      = 'HKLM\Software\Policies\Microsoft\Windows\System'
        ValueName = 'CleanupProfiles'
        Type     = 'DWord'
        Value    = $ProfileCleanupDays
        Enabled  = $EnableProfileCleanup
    }
    @{
        Setting  = 'WinRM: allow remote server management (listen on the network)'
        Key      = 'HKLM\Software\Policies\Microsoft\Windows\WinRM\Service'
        ValueName = 'AllowAutoConfig'
        Type     = 'DWord'
        Value    = 1
        Enabled  = $EnableWinRM
    }
    @{
        Setting  = 'WinRM: IPv4 listen filter'
        Key      = 'HKLM\Software\Policies\Microsoft\Windows\WinRM\Service'
        ValueName = 'IPv4Filter'
        Type     = 'String'
        Value    = '*'
        Enabled  = $EnableWinRM
    }
    @{
        Setting  = 'WinRM: IPv6 listen filter'
        Key      = 'HKLM\Software\Policies\Microsoft\Windows\WinRM\Service'
        ValueName = 'IPv6Filter'
        Type     = 'String'
        Value    = '*'
        Enabled  = $EnableWinRM
    }
    @{
        Setting  = 'WinRM: allow Basic authentication'
        Key      = 'HKLM\Software\Policies\Microsoft\Windows\WinRM\Service'
        ValueName = 'AllowBasic'
        Type     = 'DWord'
        Value    = 1
        Enabled  = $EnableWinRMBasicAuth
    }
)

# ============================================================================
# 3. FIREWALL BASELINE - written INTO the GPO, so every client gets it
# ============================================================================
$FirewallGroupName = 'AD Baseline'
# Network profile the rules apply to. 'Domain' means only networks where the
# machine can reach a domain controller - Private and Public are left unticked.
$FirewallProfile   = 'Domain'

$FirewallRules = @(
    @{ DisplayName = 'AD Baseline - Allow ICMPv4 Echo (Ping)';  Direction = 'Inbound';  Protocol = 'ICMPv4'; IcmpType = 8;   Enabled = $true }
    @{ DisplayName = 'AD Baseline - Allow ICMPv6 Echo (Ping)';  Direction = 'Inbound';  Protocol = 'ICMPv6'; IcmpType = 128; Enabled = $true }
    @{ DisplayName = 'AD Baseline - WinRM HTTP-In 5985';        Direction = 'Inbound';  Protocol = 'TCP';    Port = 5985;   Enabled = $EnableWinRM }
    @{ DisplayName = 'AD Baseline - WinRM HTTPS-In 5986';       Direction = 'Inbound';  Protocol = 'TCP';    Port = 5986;   Enabled = $EnableWinRM }
    @{ DisplayName = 'AD Baseline - WinRM Compatibility HTTP-In 80'; Direction = 'Inbound'; Protocol = 'TCP'; Port = 80;    Enabled = $EnableWinRMCompat }
    @{ DisplayName = 'AD Baseline - Remote Desktop 3389';       Direction = 'Inbound';  Protocol = 'TCP';    Port = 3389;   Enabled = $EnableRdpInbound }
    @{ DisplayName = 'AD Baseline - DNS Outbound UDP 53';       Direction = 'Outbound'; Protocol = 'UDP';    Port = 53;     Enabled = $true }
    @{ DisplayName = 'AD Baseline - DNS Outbound TCP 53';       Direction = 'Outbound'; Protocol = 'TCP';    Port = 53;     Enabled = $true }
    @{ DisplayName = 'AD Baseline - NTP Outbound UDP 123';      Direction = 'Outbound'; Protocol = 'UDP';    Port = 123;    Enabled = $true }
    @{ DisplayName = 'AD Baseline - Web Outbound HTTP 80';      Direction = 'Outbound'; Protocol = 'TCP';    Port = 80;     Enabled = $true }
    @{ DisplayName = 'AD Baseline - Web Outbound HTTPS 443';    Direction = 'Outbound'; Protocol = 'TCP';    Port = 443;    Enabled = $true }
)

# ============================================================================
# 4. CLIENT-SIDE EXTENSIONS (do not edit unless Microsoft changes them)
# ============================================================================
$RegistryMachineCSE = "[{35378EAC-683F-11D2-A89A-00C04FBBCFA2}{D02B1F72-3407-48AE-BA88-E8213C6761F1}]"
$FirewallMachineCSE = "[{35378EAC-683F-11D2-A89A-00C04FBBCFA2}{B05566AC-FE9C-4368-BE01-7A4CBB6CBA11}]"
$RegistryUserCSE    = "[{35378EAC-683F-11D2-A89A-00C04FBBCFA2}{D02B1F73-3407-48AE-BA88-E8213C6761F1}]"
$SecurityMachineCSE = "[{827D319E-6EAC-11D2-A4EA-00C04F79F83A}{803E14A0-B4FB-11D0-A0D0-00A0C90F574B}]"

$SidAuthenticatedUsers = "S-1-5-11"

# ============================================================================
# HELPERS
# ============================================================================
function Write-Step { param([string]$Text) Write-Host "`n[+] $Text" -ForegroundColor Cyan }
function Write-Ok   { param([string]$Text) Write-Host "    OK  $Text" -ForegroundColor Green }
function Write-Info { param([string]$Text) Write-Host "        $Text" -ForegroundColor Gray }
function Write-Warn { param([string]$Text) Write-Host "    !!  $Text" -ForegroundColor Yellow }

function Get-SidName {
    param([string]$Sid)
    try {
        return (New-Object System.Security.Principal.SecurityIdentifier($Sid)).Translate(
            [System.Security.Principal.NTAccount]).Value
    }
    catch { return $Sid }
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

function Select-ExistingGpo {
    $AllGpo = @(Get-GPO -All | Sort-Object DisplayName)
    if ($AllGpo.Count -eq 0) {
        throw "No GPO exists in this domain yet. Create the GPO you want in Group Policy Management, then run this script again."
    }

    Write-Host "`nExisting GPOs - pick one, this script will not create a new one:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $AllGpo.Count; $i++) {
        Write-Host (("  [{0}] {1}" -f ($i + 1), $AllGpo[$i].DisplayName))
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
    Write-Host ("  [1] Whole domain '{0}'" -f $Domain.DNSRoot)
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
        Write-Warn "'$Pick' is neither the domain root nor an existing OU. Pick a number, or paste an OU distinguished name."
    }
}

function Set-IniEntry {
    # Puts one "key=value" line into a named section of a flat line array, leaving
    # every other section and line exactly as it was. Missing sections are appended.
    param(
        [string[]]$Lines,
        [string]$Section,
        [string]$MatchPrefix,
        [string]$NewLine
    )

    $Lines = @($Lines)
    $SectionPattern = '^\s*\[' + [regex]::Escape($Section) + '\]\s*$'
    $SectionIndex = -1
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Lines[$i] -match $SectionPattern) { $SectionIndex = $i; break }
    }

    if ($SectionIndex -lt 0) {
        $Result = @($Lines) + @('', "[$Section]", $NewLine)
        return $Result
    }

    $End = $Lines.Count - 1
    for ($i = $SectionIndex + 1; $i -lt $Lines.Count; $i++) {
        if ($Lines[$i] -match '^\s*\[') { $End = $i - 1; break }
    }

    $Found = -1
    for ($i = $SectionIndex + 1; $i -le $End; $i++) {
        if ($Lines[$i].TrimStart().ToUpperInvariant().StartsWith($MatchPrefix.ToUpperInvariant())) { $Found = $i; break }
    }

    if ($Found -ge 0) {
        $Lines[$Found] = $NewLine
        return $Lines
    }

    $InsertAt = $End + 1
    $Before = if ($InsertAt -gt 0) { @($Lines[0..($InsertAt - 1)]) } else { @() }
    $After = if ($InsertAt -lt $Lines.Count) { @($Lines[$InsertAt..($Lines.Count - 1)]) } else { @() }
    return @($Before) + @($NewLine) + @($After)
}

function Read-GptTmplLines {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        # A brand new security template starts with these two headers
        return @('[Unicode]', 'Unicode=yes', '[Version]', 'signature="$CHICAGO$"', 'revision=1')
    }
    $Text = ""
    try { $Text = [System.IO.File]::ReadAllText($Path) } catch { $Text = "" }
    $Lines = @($Text -split "`r?`n")
    while ($Lines.Count -gt 0 -and [string]::IsNullOrWhiteSpace($Lines[$Lines.Count - 1])) {
        $Lines = @($Lines[0..($Lines.Count - 2)])
    }
    return @($Lines)
}

function Add-GpoClientSideExtension {
    # Registers a client-side extension on the GPO object in Active Directory so the
    # client knows it must run that extension for this policy.
    param($GpoAdsi, [string]$Property, [string]$Cse, [string]$Label)
    $Current = ""
    if ($GpoAdsi.Properties[$Property].Value) { $Current = [string]$GpoAdsi.Properties[$Property].Value }
    # .Contains() compares literally - -like would treat [ ] as wildcards
    if ($Current.Contains($Cse)) {
        Write-Info "Client-side extension already registered: $Label"
        return $false
    }
    if ($Current.Length -gt 0 -and -not $Current.EndsWith("]")) { $Current += "]" }
    $GpoAdsi.Properties[$Property].Value = $Current + $Cse
    Write-Ok "Registered client-side extension: $Label"
    return $true
}

$Script:ResultCounts = @{ Changed = 0; Unchanged = 0; Failed = 0 }
$Script:MachineVersionBump = 0

$StartTime = Get-Date

try {
    # ========================================================================
    # 0. RESOLVE DOMAIN, GPO AND LINK TARGET
    # ========================================================================
    if (-not [System.Environment]::UserInteractive) {
        throw "This script asks questions, so it must run in an interactive PowerShell window."
    }

    try { $AdDomain = Get-ADDomain -ErrorAction Stop }
    catch {
        throw "Cannot read the Active Directory domain from this machine. Run the script as a Domain Admin on a domain-joined server. Reason: $($_.Exception.Message)"
    }

    $DomainDN    = [string]$AdDomain.DistinguishedName
    $DomainFQDN  = [string]$AdDomain.DNSRoot
    $DomainNetBiosName = [string]$AdDomain.NetBIOSName
    if ([string]::IsNullOrWhiteSpace($DomainNetBIOS)) { $DomainNetBIOS = $DomainNetBiosName }
    if ([string]::IsNullOrWhiteSpace($DomainNetBIOS)) { $DomainNetBIOS = $DomainFQDN.Split('.')[0] }

    if ([string]::IsNullOrWhiteSpace($TargetOU)) {
        $TargetOU = Select-GpoLinkTarget -Domain $AdDomain
    }
    $TargetOU = $TargetOU.Trim()
    if (-not (Test-GpoLinkTarget -DistinguishedName $TargetOU -DomainDN $DomainDN)) {
        Write-Warn "'$TargetOU' is not the domain root and not an OU, so Windows cannot link a GPO there."
        $TargetOU = Select-GpoLinkTarget -Domain $AdDomain
    }

    Write-Info "Detected: domain '$DomainFQDN' (NetBIOS '$DomainNetBIOS')"

    Write-Step "1/6  Existing GPO"
    if ([string]::IsNullOrWhiteSpace($GPOName)) { $GPOName = Select-ExistingGpo }

    $GPO = Get-GPO -Name $GPOName -ErrorAction SilentlyContinue
    if (-not $GPO) {
        throw "There is no GPO named '$GPOName' in '$DomainFQDN'. This script never creates one - create it in Group Policy Management first, then run this again."
    }
    Write-Ok "Using existing GPO '$GPOName' (GUID $($GPO.Id))."

    $GPOGuid   = $GPO.Id.ToString("B").ToUpper()
    $SysvolGpo = "\\$DomainFQDN\SYSVOL\$DomainFQDN\Policies\$GPOGuid"
    Write-Info "GPO folder: $SysvolGpo"

    # ---- link the GPO to the chosen scope --------------------------------
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
        if (Read-YesNo -Prompt 'Link it now?') {
            try {
                New-GPLink -Name $GPOName -Target $TargetOU -LinkEnabled Yes -ErrorAction Stop | Out-Null
                Write-Ok "Linked GPO to '$TargetOU'."
            }
            catch {
                if ($_.Exception.Message -match 'already (exists|linked|in place)') { Write-Ok "GPO already linked to '$TargetOU'." }
                else { throw }
            }
        }
        else {
            Write-Warn "Not linked - the settings are still saved in the GPO, they just apply nowhere yet."
        }
    }
    else {
        Write-Ok "GPO already linked to '$TargetOU'."
    }

    # ========================================================================
    # 2. SHOW THE PLAN, THEN ASK
    # ========================================================================
    Write-Step "2/6  Settings this script will write into '$GPOName'"
    Write-Host "  User Configuration:" -ForegroundColor White
    foreach ($S in $UserPolicySettings) {
        if ($S.Enabled) { Write-Info ("[set]   {0}  ({1}\{2}={3})" -f $S.Setting, $S.Key, $S.ValueName, $S.Value) }
        else { Write-Info ("[skip]  {0}" -f $S.Setting) }
    }
    Write-Host "  Computer Configuration:" -ForegroundColor White
    foreach ($S in $ComputerPolicySettings) {
        if ($S.Enabled) { Write-Info ("[set]   {0}  ({1}\{2}={3})" -f $S.Setting, $S.Key, $S.ValueName, $S.Value) }
        else { Write-Info ("[skip]  {0}" -f $S.Setting) }
    }
    Write-Host "  Security Policy (GptTmpl.inf):" -ForegroundColor White
    if ($DisableCredentialCache) { Write-Info "[set]   Interactive logon: number of previous logons to cache = 0" }
    if ($EnableWinRM) { Write-Info "[set]   Service 'WinRM' start type = Automatic" }
    if ($EnableFirewallRules) { Write-Info ("[set]   {0} firewall rules for group '{1}' (domain profile only)" -f @($FirewallRules | Where-Object { $_.Enabled }).Count, $FirewallGroupName) }

    Write-Host ""
    if (-not (Read-YesNo -Prompt "Apply these settings to the GPO '$GPOName'?")) {
        throw "Stopped: nothing was written to the GPO."
    }

    # ========================================================================
    # 3. REGISTRY POLICIES (registry.pol through the Group Policy cmdlets)
    # ========================================================================
    Write-Step "3/6  Registry policies"

    function Set-OnePolicySetting {
        param($Definition)
        try {
            $Current = Get-GPRegistryValue -Name $GPOName -Key $Definition.Key -ValueName $Definition.ValueName -ErrorAction SilentlyContinue
            if ($Current -and $Current.PolicyType -eq 'Registry' -and ([string]$Current.Value) -eq ([string]$Definition.Value)) {
                Write-Info "unchanged: $($Definition.Setting)"
                $Script:ResultCounts.Unchanged++
                return
            }
            Set-GPRegistryValue -Name $GPOName -Key $Definition.Key -ValueName $Definition.ValueName `
                                -Type $Definition.Type -Value $Definition.Value | Out-Null
            Write-Ok "set: $($Definition.Setting)  ($($Definition.ValueName)=$($Definition.Value))"
            $Script:ResultCounts.Changed++
        }
        catch {
            Write-Warn "FAILED: $($Definition.Setting) - $($_.Exception.Message)"
            $Script:ResultCounts.Failed++
        }
    }

    foreach ($S in $UserPolicySettings) {
        if ($S.Enabled) { Set-OnePolicySetting -Definition $S }
        else { Write-Info "skipped by config: $($S.Setting)" }
    }
    foreach ($S in $ComputerPolicySettings) {
        if ($S.Enabled) { Set-OnePolicySetting -Definition $S }
        else { Write-Info "skipped by config: $($S.Setting)" }
    }

    # ========================================================================
    # 4. SECURITY POLICY ENTRIES (merge into GptTmpl.inf, keep everything else)
    # ========================================================================
    Write-Step "4/6  Security policy entries in GptTmpl.inf"
    $GptTmplPath = Join-Path (Join-Path $SysvolGpo "Machine") "Microsoft\Windows NT\SecEdit\GptTmpl.inf"
    $SecEditDir  = Split-Path -Parent $GptTmplPath
    $GptChanged  = $false

    if (($DisableCredentialCache -or $EnableWinRM)) {
        $Lines = Read-GptTmplLines -Path $GptTmplPath
        $OriginalText = ($Lines -join "`r`n")

        if ($DisableCredentialCache) {
            # 1 = REG_SZ, ""0"" is the cached logon count security policy writes
            $Lines = Set-IniEntry -Lines $Lines -Section 'Registry Values' `
                -MatchPrefix 'MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\CachedLogonsCount=' `
                -NewLine 'MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\CachedLogonsCount=1,"0"'
        }
        if ($EnableWinRM) {
            # Startup mode: 2 = Automatic, 3 = Manual, 4 = Disabled
            $Lines = Set-IniEntry -Lines $Lines -Section 'Service General Setting' `
                -MatchPrefix 'WinRM,' -NewLine 'WinRM,2,A'
        }

        $NewText = ($Lines -join "`r`n")
        if ($NewText -ne $OriginalText) {
            if (-not (Test-Path -LiteralPath $SecEditDir)) { New-Item -Path $SecEditDir -ItemType Directory -Force | Out-Null }
            [System.IO.File]::WriteAllText($GptTmplPath, ($NewText + "`r`n"),
                (New-Object System.Text.UnicodeEncoding($false, $true)))
            $GptChanged = $true
            Write-Ok "GptTmpl.inf updated: CachedLogonsCount$(if ($DisableCredentialCache) { '=0' }); WinRM service$(if ($EnableWinRM) { ' = Automatic' })."
        }
        else {
            Write-Info "GptTmpl.inf already had these entries - left untouched."
        }
    }
    else {
        Write-Info "Both security entries are switched off in the configuration - nothing to merge."
    }

    # ========================================================================
    # 5. FIREWALL RULES INSIDE THE GPO
    # ========================================================================
    Write-Step "5/6  Firewall rules in the GPO"
    $FirewallChanged = 0
    if ($EnableFirewallRules) {
        # The firewall cmdlets take a policy store written as '<FQDN domain>\<GPO name>'
        # (or the GPO GUID). The NetBIOS name is refused with "The parameter is
        # incorrect.", and Open-NetGPO has no -Name parameter - it takes -PolicyStore too.
        $GpoStore = "$DomainFQDN\$GPOName"
        $Session  = $null
        $GpoParam = @{}
        try {
            $Session  = Open-NetGPO -PolicyStore $GpoStore -ErrorAction Stop
            $GpoParam = @{ GPOSession = $Session }
        }
        catch {
            Write-Warn "Open-NetGPO failed ($($_.Exception.Message)) - writing straight to the policy store '$GpoStore'."
            $GpoParam = @{ PolicyStore = $GpoStore }
        }

        try {
            foreach ($Rule in $FirewallRules) {
                if (-not $Rule.Enabled) {
                    Write-Info "skipped by config: $($Rule.DisplayName)"
                    continue
                }

                $Present = $null
                try { $Present = Get-NetFirewallRule @GpoParam -DisplayName $Rule.DisplayName -ErrorAction Stop }
                catch { $Present = $null }

                if ($Present) {
                    # A rule left by an earlier run may still carry the old profile
                    # (Any = Domain + Private + Public). Narrow it to Domain only.
                    $NowProfile = @($Present.Profile) -join ','
                    if ($NowProfile -ne $FirewallProfile) {
                        try {
                            Set-NetFirewallRule @GpoParam -DisplayName $Rule.DisplayName -Profile $FirewallProfile -ErrorAction Stop | Out-Null
                            Write-Ok "profile -> $FirewallProfile : $($Rule.DisplayName)"
                            $FirewallChanged++
                            $Script:ResultCounts.Changed++
                        }
                        catch {
                            Write-Warn "FAILED to set profile: $($Rule.DisplayName) - $($_.Exception.Message)"
                            $Script:ResultCounts.Failed++
                        }
                    }
                    else {
                        Write-Info "already in the GPO: $($Rule.DisplayName)"
                        $Script:ResultCounts.Unchanged++
                    }
                    continue
                }

                $Params = @{
                    DisplayName = $Rule.DisplayName
                    Direction   = $Rule.Direction
                    Action      = 'Allow'
                    Profile     = $FirewallProfile
                    Enabled     = 'True'
                    Group       = $FirewallGroupName
                }
                if ($Rule.Protocol -eq 'ICMPv4' -or $Rule.Protocol -eq 'ICMPv6') {
                    $Params['Protocol'] = $Rule.Protocol
                    $Params['IcmpType'] = $Rule.IcmpType
                }
                else {
                    $Params['Protocol'] = $Rule.Protocol
                    $Params['LocalPort'] = $Rule.Port
                }
                foreach ($Key in $GpoParam.Keys) { $Params[$Key] = $GpoParam[$Key] }

                # One bad rule must not stop the rest of the baseline.
                try {
                    New-NetFirewallRule @Params -ErrorAction Stop | Out-Null
                    Write-Ok "created: $($Rule.DisplayName)"
                    $FirewallChanged++
                    $Script:ResultCounts.Changed++
                }
                catch {
                    Write-Warn "FAILED: $($Rule.DisplayName) - $($_.Exception.Message)"
                    $Script:ResultCounts.Failed++
                }
            }

            if ($Session -and $FirewallChanged -gt 0) {
                Save-NetGPO -GPOSession $Session -ErrorAction Stop
                Write-Ok "Saved $FirewallChanged new firewall rule(s) into the GPO."
            }
        }
        catch {
            Write-Warn "Firewall rules could not all be written: $($_.Exception.Message)"
            $Script:ResultCounts.Failed++
        }
        finally {
            # Closing is best effort: not every build ships a close cmdlet, and a
            # session left open costs nothing once the script exits.
            if ($Session) {
                if (Get-Command -Name 'Close-NetGPO' -ErrorAction SilentlyContinue) {
                    try { Close-NetGPO -GPOSession $Session } catch { }
                }
                $Session = $null
            }
        }
        # The firewall cmdlets bump the GPO version in Active Directory themselves,
        # so nothing is added here - step 6 only brings GPT.INI back in line.
    }
    else {
        Write-Info "Firewall baseline switched off in the configuration - nothing written."
    }

    # ========================================================================
    # 6. CSE REGISTRATION, VERSION BUMP AND GPT.INI
    # ========================================================================
    Write-Step "6/6  Client-side extensions, GPO version and SYSVOL"
    $GpoDn = "CN=$GPOGuid,CN=Policies,CN=System,$DomainDN"
    $GpoAdsi = [ADSI]"LDAP://$GpoDn"

    $CseChanged = $false
    if ($FirewallChanged -gt 0) {
        if (Add-GpoClientSideExtension -GpoAdsi $GpoAdsi -Property 'gPCMachineExtensionNames' `
                -Cse $FirewallMachineCSE -Label 'Windows Firewall policy') { $CseChanged = $true; $Script:MachineVersionBump += 1 }
    }
    if ($GptChanged) {
        if (Add-GpoClientSideExtension -GpoAdsi $GpoAdsi -Property 'gPCMachineExtensionNames' `
                -Cse $SecurityMachineCSE -Label 'Security policy (GptTmpl.inf)') { $CseChanged = $true }
        $Script:MachineVersionBump += 1
    }
    # Registry policies also need the Registry extension; add it only if it is missing
    # (Set-GPRegistryValue normally registers it itself, this is the safety net).
    $HasUserRegistry = $false
    foreach ($S in $UserPolicySettings) { if ($S.Enabled) { $HasUserRegistry = $true } }
    $HasMachineRegistry = $false
    foreach ($S in $ComputerPolicySettings) { if ($S.Enabled) { $HasMachineRegistry = $true } }
    if ($HasMachineRegistry) {
        if (Add-GpoClientSideExtension -GpoAdsi $GpoAdsi -Property 'gPCMachineExtensionNames' `
                -Cse $RegistryMachineCSE -Label 'Registry policy (machine)') { $CseChanged = $true; $Script:MachineVersionBump += 1 }
    }
    if ($HasUserRegistry) {
        if (Add-GpoClientSideExtension -GpoAdsi $GpoAdsi -Property 'gPCUserExtensionNames' `
                -Cse $RegistryUserCSE -Label 'Registry policy (user)') { $CseChanged = $true; $Script:MachineVersionBump += 65536 }
    }

    if ($CseChanged -or $Script:MachineVersionBump -gt 0 -or $FirewallChanged -gt 0) {
        # Read the version here, after the firewall writes, so the number the firewall
        # cmdlets already bumped is kept instead of being counted twice.
        $OldVersion = [int]$GpoAdsi.Properties["versionNumber"].Value
        $NewVersion = $OldVersion + $Script:MachineVersionBump
        if ($NewVersion -ne $OldVersion) {
            $GpoAdsi.Properties["versionNumber"].Value = $NewVersion
            $GpoAdsi.CommitChanges()
        }
        Write-Ok "GPO version: $OldVersion -> $NewVersion (extra bumps this run: $Script:MachineVersionBump; firewall rules written: $FirewallChanged)."

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

    # ---- wait for SYSVOL replication -------------------------------------
    if ($SysvolWaitSeconds -gt 0 -and ($Script:MachineVersionBump -gt 0 -or $FirewallChanged -gt 0)) {
        $DcHosts = @(Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName)
        if ($DcHosts.Count -gt 1) {
            Write-Info "Waiting up to $SysvolWaitSeconds s for SYSVOL to reach $($DcHosts.Count) DCs..."
            $Deadline = (Get-Date).AddSeconds($SysvolWaitSeconds)
            $Pending = @($DcHosts | Where-Object {
                -not (Test-Path -LiteralPath "\\$_\SYSVOL\$DomainFQDN\Policies\$GPOGuid\GPT.INI")
            })
            while ($Pending.Count -gt 0 -and (Get-Date) -lt $Deadline) {
                Start-Sleep -Seconds 5
                $Pending = @($Pending | Where-Object {
                    -not (Test-Path -LiteralPath "\\$_\SYSVOL\$DomainFQDN\Policies\$GPOGuid\GPT.INI")
                })
            }
            if ($Pending.Count -eq 0) { Write-Ok "SYSVOL settings copied to all domain controllers." }
            else { Write-Warn "Still copying to: $($Pending -join ', '). Check with repadmin /showrepl later." }
        }
    }

    # ---- Done ------------------------------------------------------------
    $AuthUsersName = Get-SidName -Sid $SidAuthenticatedUsers
    Write-Host "`n=== Finished in $([int]((Get-Date) - $StartTime).TotalSeconds)s ===" -ForegroundColor Green
    Write-Host "GPO           : $GPOName ($GPOGuid)"
    Write-Host "Applies to    : $TargetOU"
    Write-Host "Changed       : $($Script:ResultCounts.Changed)"
    Write-Host "Already set   : $($Script:ResultCounts.Unchanged)"
    Write-Host "Failures      : $($Script:ResultCounts.Failed)"
    Write-Host "`nHow to verify:" -ForegroundColor Cyan
    Write-Info "1. gpupdate /force on a client, then sign out and back in."
    Write-Info "2. gpresult /scope user /h report.html  -> regedit should say the policy blocks it."
    Write-Info "3. gpresult /scope computer /h report.html -> persona/wallpaper, Store, Offline Files, WinRM and firewall rules."
    Write-Info "4. 'Get-NetFirewallRule -DisplayName ''AD Baseline*''' on the client lists the rules from the GPO."
    Write-Info "5. 'Test-WSMan <client>' or 'winrm qc' from an admin PC proves WinRM; 'ping <client>' proves ICMP."
    Write-Info "6. Security policy: run 'net accounts' - 'Logon hours / logon limit' section shows it; walk 'secpol.msc' for the cache setting."
    if ($Script:ResultCounts.Failed -gt 0) {
        Write-Warn "$($Script:ResultCounts.Failed) setting(s) could not be written - see the !! lines above."
    }
}
catch {
    Write-Host "`n=== Deployment FAILED ===" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "At line: $($_.InvocationInfo.ScriptLineNumber)  ->  $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkYellow
    exit 1
}
