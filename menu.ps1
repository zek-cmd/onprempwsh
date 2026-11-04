#Requires -Version 5.1
<#
.SYNOPSIS
    One menu that runs any of the on-premises AD scripts in this folder.

.DESCRIPTION
    Shows the scripts in this folder as a short numbered menu, grouped by what they
    do (user accounts, moving users, passwords, reports and cleanup, software and
    computers, DFS, just-in-time access).

    Before a script is started the menu shows what it does and the exact command
    line, then asks a y/n confirmation. Scripts that can delete or overwrite data
    ask you to type YES instead of a plain y.

    -WhatIf puts the menu in dry-run mode: the y/n confirmation is still asked, and
the script you pick is still started - but the menu adds -WhatIf to its command
line, so the script runs in its own preview mode and reports what it WOULD do
without changing anything. Scripts that have no -WhatIf switch are not started in
this mode, because running them would make real changes. While the menu is open,
"W" turns dry-run mode on and off.

    Each script runs in its own PowerShell session inside this console window, so a
    script that finishes or fails comes back to the menu.

.PARAMETER WhatIf
    Dry run. Start the chosen script with -WhatIf added, so it previews instead of
    changing anything. Scripts without a -WhatIf switch are not started.

.EXAMPLE
    .\menu.ps1

.EXAMPLE
    .\menu.ps1 -WhatIf

.EXAMPLE
    .\menu.ps1 -WhatIf:$false
#>

param(
    [switch]$WhatIf
)

$ErrorActionPreference = 'Continue'

$ScriptRoot = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptRoot)) {
    $ScriptRoot = (Get-Location).Path
}

$ReportFolder = Join-Path $ScriptRoot 'Output Report'

# ============================================
# HELPERS
# ============================================

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = 'White'
    )
    Write-Host $Message -ForegroundColor $Color
}

function Write-Banner {
    param([string]$Title)
    Write-ColorOutput '================================================' 'Cyan'
    Write-ColorOutput ("   {0}" -f $Title) 'Yellow'
    Write-ColorOutput '================================================' 'Cyan'
}

function Get-RiskText {
    param([string]$Risk)
    switch ($Risk) {
        'Danger' { return 'DANGER - can delete or overwrite data' }
        'Write' { return 'changes data' }
        default { return 'read only' }
    }
}

function Get-RiskShort {
    param([string]$Risk)
    switch ($Risk) {
        'Danger' { return 'DANGER' }
        'Write' { return 'changes' }
        default { return 'read' }
    }
}

function Get-RiskColor {
    param([string]$Risk)
    switch ($Risk) {
        'Danger' { return 'Red' }
        'Write' { return 'Yellow' }
        default { return 'Gray' }
    }
}

# The PowerShell program to start a script with. Scripts marked as needing 7.0
# use pwsh.exe; the rest use Windows PowerShell 5.1, which is what they were
# written for. Returns $null when a 7.0 script is picked and pwsh is missing.
function Get-CurrentHostExe {
    try {
        $exe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if ($exe -match '(?i)(powershell|pwsh)\.exe$') { return $exe }
    }
    catch { }

    if ($PSVersionTable.PSVersion.Major -ge 7) { return 'pwsh.exe' }
    return 'powershell.exe'
}

function Resolve-ScriptHost {
    param([switch]$Needs7)

    $isSeven = $PSVersionTable.PSVersion.Major -ge 7

    if ($Needs7) {
        if ($isSeven) { return (Get-CurrentHostExe) }
        $pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
        if ($pwsh) { return $pwsh.Source }
        return $null
    }

    $ps5 = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($ps5) { return $ps5.Source }
    return (Get-CurrentHostExe)
}

# Split the typed arguments into a list, keeping "quoted values" together.
function ConvertFrom-ArgumentLine {
    param([string]$Line)

    $list = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($Line, '(?:"([^"]*)"|(\S+))')) {
        if ($m.Groups[1].Success) {
            [void]$list.Add($m.Groups[1].Value)
        }
        else {
            [void]$list.Add($m.Groups[2].Value)
        }
    }
    return $list.ToArray()
}

function Get-CommandText {
    param(
        [string]$HostExe,
        [string]$ScriptPath,
        [string[]]$Arguments
    )

    $text = '{0} -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $HostExe, $ScriptPath
    foreach ($a in @($Arguments)) {
        if ([string]::IsNullOrEmpty($a)) { continue }
        if ($a -match '\s') { $text += ' "' + $a + '"' }
        else { $text += ' ' + $a }
    }
    return $text
}

function New-MenuItem {
    param(
        [string]$Group,
        [string]$File,
        [string]$About,
        [string]$Risk = 'Read',
        [switch]$Needs7,
        [switch]$SupportsWhatIf
    )

    [PSCustomObject]@{
        Group           = $Group
        File            = $File
        About           = $About
        Risk            = $Risk
        Needs7          = [bool]$Needs7
        SupportsWhatIf  = [bool]$SupportsWhatIf
    }
}

# ============================================
# SCRIPT LIST
# ============================================

$Catalog = @(
    New-MenuItem -Group 'User accounts' -File 'USERCREATE.ps1' -About 'Create users from a CSV (FullName column)' -Risk Write
    New-MenuItem -Group 'User accounts' -File 'USERCREATE2.ps1' -About 'Create users from a CSV - alternate version' -Risk Write
    New-MenuItem -Group 'User accounts' -File 'nameconvert.ps1' -About 'Convert LastName, FirstName to FirstName LastName' -Risk Read

    New-MenuItem -Group 'Moving users' -File 'MOVEUSER.ps1' -About 'Move all users from one OU to another' -Risk Write
    New-MenuItem -Group 'Moving users' -File 'MOVEUSER2.ps1' -About 'Move users - alternate version' -Risk Write
    New-MenuItem -Group 'Moving users' -File 'MOVEUSER3.ps1' -About 'Move users - alternate version' -Risk Write

    New-MenuItem -Group 'Passwords' -File 'RESETPASSWORD.ps1' -About 'Reset passwords for all users in an OU' -Risk Write
    New-MenuItem -Group 'Passwords' -File 'RESETPASSWORD2.ps1' -About 'Reset passwords - alternate version' -Risk Write

    New-MenuItem -Group 'Reports and cleanup' -File 'INACTIVITY.ps1' -About 'Find inactive accounts and their folders' -Risk Write -SupportsWhatIf
    New-MenuItem -Group 'Reports and cleanup' -File 'Find-Duplicate-AD-Names.ps1' -About 'Find duplicate or close display names' -Risk Read
    New-MenuItem -Group 'Reports and cleanup' -File 'folderbased\folderlookup.ps1' -About 'Look up each user redirected folders' -Risk Read
    New-MenuItem -Group 'Reports and cleanup' -File 'folderbased\folderlookuponlogon.ps1' -About 'Folder audit report on logon' -Risk Read
    New-MenuItem -Group 'Reports and cleanup' -File 'folderbased\extract.ps1' -About 'Extract .xlsx/.docx files from user folders' -Risk Write
    New-MenuItem -Group 'Reports and cleanup' -File 'folderbased\delete.ps1' -About 'Delete the extracted .xlsx/.docx files' -Risk Danger
    New-MenuItem -Group 'Reports and cleanup' -File 'folderbased\folderredirection.ps1' -About 'Reset folder ownership and permissions' -Risk Danger

    New-MenuItem -Group 'Software and computers' -File 'SOFTWAREDEPLOY.ps1' -About 'Create and share a GPO software folder' -Risk Write
    New-MenuItem -Group 'Software and computers' -File 'SHUTDOWN.ps1' -About 'Shutdown, restart or log off computers' -Risk Write -Needs7
    New-MenuItem -Group 'Software and computers' -File 'message.ps1' -About 'Send a message to remote computers' -Risk Write

    New-MenuItem -Group 'DFS namespaces' -File 'folderbased\DFS_script\install_dfs.ps1' -About 'Install DFS Namespace and its tools' -Risk Write -SupportsWhatIf
    New-MenuItem -Group 'DFS namespaces' -File 'folderbased\DFS_script\create_dfsn.ps1' -About 'Create a DFS namespace' -Risk Write
    New-MenuItem -Group 'DFS namespaces' -File 'folderbased\DFS_script\check_dfs.ps1' -About 'Show the existing DFS namespaces' -Risk Read

    New-MenuItem -Group 'Just-in-time access' -File 'JIT\Start-JITServer.ps1' -About 'Start the local JIT admin web server' -Risk Write -Needs7
)

$Groups = @($Catalog | Group-Object -Property Group)

# ============================================
# RUN ONE SCRIPT (with WhatIf and y/n confirmation)
# ============================================

function Invoke-MenuScript {
    param($Item)

    $leaf = Split-Path -Leaf $Item.File
    $fullPath = Join-Path $ScriptRoot $Item.File

    Write-Host ''
    Write-ColorOutput '------------------------------------------------' 'Cyan'
    Write-ColorOutput ("Script : {0}" -f $leaf) 'White'
    Write-ColorOutput ("File   : {0}" -f $fullPath) 'Gray'
    Write-ColorOutput ("Does   : {0}" -f $Item.About) 'White'
    Write-ColorOutput ("Level  : {0}" -f (Get-RiskText -Risk $Item.Risk)) (Get-RiskColor -Risk $Item.Risk)
    Write-ColorOutput '------------------------------------------------' 'Cyan'

    if (-not (Test-Path -LiteralPath $fullPath)) {
        Write-ColorOutput "`nScript not found. Check the file name in this folder:" 'Red'
        Write-ColorOutput "  $fullPath" 'Gray'
        return
    }

    $hostExe = Resolve-ScriptHost -Needs7:$Item.Needs7
    if ([string]::IsNullOrEmpty($hostExe)) {
        Write-ColorOutput "`nThis script needs PowerShell 7.0 (pwsh.exe), which is not installed." 'Red'
        Write-ColorOutput 'Install PowerShell 7 and try again, or run this script yourself from PowerShell 7.' 'Gray'
        return
    }

    $extra = Read-Host "`nExtra arguments for the script (press Enter for none)"
    $argList = @()
    if (-not [string]::IsNullOrWhiteSpace($extra)) {
        $argList = @(ConvertFrom-ArgumentLine -Line $extra)
    }

    $commandText = Get-CommandText -HostExe $hostExe -ScriptPath $fullPath -Arguments $argList

    Write-ColorOutput "`nCommand that will be used:" 'Cyan'
    Write-ColorOutput "  $commandText" 'Gray'

    # ---------- y/n confirmation first ----------
    $answer = Read-Host "`nRun $leaf now? (y/n)"
    if ($answer -notmatch '^(y|yes)$') {
        Write-ColorOutput 'Cancelled - nothing was run.' 'Gray'
        return
    }

    if ($Item.Risk -eq 'Danger') {
        $typed = Read-Host "This script can delete or overwrite data. Type YES to continue"
        if ($typed -cne 'YES') {
            Write-ColorOutput 'Cancelled - nothing was run.' 'Gray'
            return
        }
    }

    # ---------- WhatIf: start it, but in its own preview mode ----------
    $runArgs = $argList
    if ($script:DryRun) {

        # A read-only script changes nothing either way. A script that changes data
        # is only safe to start when it understands -WhatIf itself.
        if ($Item.Risk -ne 'Read' -and -not $Item.SupportsWhatIf) {
            Write-ColorOutput "`n[WHATIF] Dry-run mode is ON and $leaf has no -WhatIf switch." 'Yellow'
            Write-ColorOutput '[WHATIF] Starting it would make real changes, so it was not started.' 'Yellow'
            Write-ColorOutput '[WHATIF] Press W in the menu to turn dry-run mode off first.' 'Gray'
            return
        }

        if (@($runArgs) -notcontains '-WhatIf') {
            $runArgs = @(@($runArgs) + '-WhatIf')
        }

        $previewText = Get-CommandText -HostExe $hostExe -ScriptPath $fullPath -Arguments $runArgs
        Write-ColorOutput "`n[WHATIF] Dry-run mode is ON - starting the script in its preview mode." 'Yellow'
        Write-ColorOutput "[WHATIF] $previewText" 'Yellow'
        if ($Item.Risk -eq 'Read') {
            Write-ColorOutput '[WHATIF] This script only reads, so the result is the same as a normal run.' 'Gray'
        }
        else {
            Write-ColorOutput '[WHATIF] The script will report what it WOULD do and change nothing.' 'Gray'
        }
    }

    Write-ColorOutput ("`nStarting {0} ..." -f $leaf) 'Green'

    Push-Location -LiteralPath $ScriptRoot
    try {
        & $hostExe -NoProfile -ExecutionPolicy Bypass -File $fullPath @runArgs
    }
    catch {
        Write-ColorOutput ("The script could not be started: {0}" -f $_.Exception.Message) 'Red'
    }
    finally {
        Pop-Location
    }

    Write-ColorOutput ("`n{0} finished - back to the menu." -f $leaf) 'Green'
}

# ============================================
# CATEGORY SCREEN
# ============================================

function Invoke-CategoryScreen {
    param(
        [string]$GroupName,
        $Items
    )

    $items = @($Items)

    while ($true) {
        Clear-Host
        Write-Banner $GroupName
        Write-ColorOutput ("   Mode: {0}" -f (Get-ModeText)) (Get-ModeColor)
        Write-Host ''

        for ($i = 1; $i -le $items.Count; $i++) {
            $item = $items[$i - 1]
            $name = Split-Path -Leaf $item.File
            if ($name.Length -gt 30) { $name = $name.Substring(0, 27) + '...' }
            $about = $item.About
            if ($about.Length -gt 42) { $about = $about.Substring(0, 39) + '...' }

            $line = '  {0,2}) {1,-30}  {2,-42}  [{3}]' -f $i, $name, $about, (Get-RiskShort -Risk $item.Risk)
            Write-ColorOutput $line (Get-RiskColor -Risk $item.Risk)

            if ($script:DryRun -and $item.Risk -ne 'Read' -and -not $item.SupportsWhatIf) {
                Write-ColorOutput '      (no preview mode - it will not start while WhatIf is on)' 'DarkGray'
            }
        }

        Write-Host ''
        Write-ColorOutput '------------------------------------------------' 'Cyan'
        Write-ColorOutput '   B) Back to the main menu' 'Cyan'

        $choice = Read-Host "`nChoose a number or B"
        if ([string]::IsNullOrWhiteSpace($choice)) { continue }

        $c = $choice.Trim()

        if ($c -match '^(b|back|q)$') { return }
        if ($c -match '^(w|whatif)$') {
            $script:DryRun = -not $script:DryRun
            continue
        }

        $number = 0
        if ([int]::TryParse($c, [ref]$number) -and $number -ge 1 -and $number -le $items.Count) {
            Invoke-MenuScript -Item $items[$number - 1]
            Read-Host 'Press Enter to go back to the list'
            continue
        }

        Write-ColorOutput "`n'$c' is not one of the choices. Try a number from the list, W or B." 'Red'
        Start-Sleep -Milliseconds 1200
    }
}

# ============================================
# MODE AND ENVIRONMENT TEXT
# ============================================

function Get-ModeText {
    if ($script:DryRun) { return 'WHATIF - scripts start in their own preview mode' }
    return 'LIVE - each script asks y/n before it starts'
}

function Get-ModeColor {
    if ($script:DryRun) { return 'Yellow' }
    return 'Green'
}

function Open-ReportFolder {
    if (Test-Path -LiteralPath $ReportFolder) {
        Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$ReportFolder`""
        Write-ColorOutput "`nOpened $ReportFolder" 'Gray'
    }
    else {
        Write-ColorOutput "`nNo report folder yet. It is created the first time a script saves a report." 'Yellow'
        Write-ColorOutput "  $ReportFolder" 'Gray'
    }
}

# ============================================
# MAIN MENU
# ============================================

$domainName = $env:USERDNSDOMAIN
if ([string]::IsNullOrEmpty($domainName)) { $domainName = $env:USERDOMAIN }
if ([string]::IsNullOrEmpty($domainName)) { $domainName = '(not domain-joined)' }

try {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
catch {
    $isAdmin = $false
}

$script:DryRun = [bool]$WhatIf

while ($true) {
    Clear-Host
    Write-Banner 'ON-PREM AD SCRIPT MENU'
    Write-ColorOutput ("   Domain : {0}" -f $domainName) 'Gray'
    Write-ColorOutput ("   Folder : {0}" -f $ScriptRoot) 'Gray'
    if ($isAdmin) { Write-ColorOutput '   Rights : Administrator' 'Gray' }
    else { Write-ColorOutput '   Rights : standard user (some scripts will fail)' 'Yellow' }
    Write-ColorOutput ("   Mode   : {0}" -f (Get-ModeText)) (Get-ModeColor)
    Write-Host ''

    for ($i = 1; $i -le $Groups.Count; $i++) {
        $group = $Groups[$i - 1]
        $count = @($group.Group).Count
        $label = if ($count -eq 1) { '1 script' } else { "$count scripts" }
        $line = '  {0,2}) {1,-24} ({2})' -f $i, $group.Name, $label
        Write-ColorOutput $line 'White'
    }

    Write-Host ''
    Write-ColorOutput '------------------------------------------------' 'Cyan'
    Write-ColorOutput '   W) Turn dry-run (WhatIf) mode on or off' 'Cyan'
    Write-ColorOutput '   R) Open the report folder' 'Cyan'
    Write-ColorOutput '   Q) Quit' 'Cyan'

    $choice = Read-Host "`nChoose a category, W, R or Q"
    if ([string]::IsNullOrWhiteSpace($choice)) { continue }

    $c = $choice.Trim()

    if ($c -match '^(q|quit|exit)$') { break }

    if ($c -match '^(w|whatif)$') {
        $script:DryRun = -not $script:DryRun
        continue
    }

    if ($c -match '^(r|report)$') {
        Open-ReportFolder
        Read-Host 'Press Enter to go back to the menu'
        continue
    }

    $number = 0
    if ([int]::TryParse($c, [ref]$number) -and $number -ge 1 -and $number -le $Groups.Count) {
        $group = $Groups[$number - 1]
        Invoke-CategoryScreen -GroupName $group.Name -Items @($group.Group)
        continue
    }

    Write-ColorOutput "`n'$c' is not one of the choices. Try a category number, W, R or Q." 'Red'
    Start-Sleep -Milliseconds 1200
}

# ============================================
# EXIT
# ============================================

Clear-Host
Write-Banner 'MENU CLOSED'
Write-ColorOutput ("   Closed at: {0}" -f (Get-Date)) 'Cyan'
Write-ColorOutput '   No script was left running. Nothing was changed by this menu.' 'Green'
Write-ColorOutput '================================================' 'Cyan'
