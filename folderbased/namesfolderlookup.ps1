#Requires -Version 5.1
<#
.SYNOPSIS
    Compares a names file (names.csv) with the per-user folders inside a folder
    redirection root, then copies, moves, renames, deletes or exports them.

.DESCRIPTION
    The folder redirection script creates one folder per user under the redirect
    root (the folder named after %USERNAME%). This script:

      1. asks for the names file (names.csv, or any CSV with a Username and/or
         FullName column) in a normal Windows Open dialog,
      2. asks for the redirect root folder in a normal Windows folder dialog,
      3. reads the folders directly inside that root and compares each folder
         name with the names in the CSV,
      4. prints a report in three parts: folders that match a name, folders that
         match nothing, and names that have no folder at all,
      5. offers the operations: copy, move, rename to FullName, delete, export
         the report to CSV, rescan.

    Destination naming:
      A copied or moved folder lands under the full name from the CSV - "First
      Name Middle Initial Last Name" - instead of the AD username. A name written
      the other way round ("Surname, First Name") is swapped into that basis.
      Press N in the menu to switch between the full name and the username.

    Folder matching:
      Matched  - the folder name equals a Username (or FullName) from the CSV.
      Check    - the names look alike once dots, spaces and capitals are ignored,
                 so it is probably the right person - look before you act.
      No match - no name in the CSV fits this folder (leftover / old folder).

    Nothing is changed until an operation is picked. Every operation asks which
    folders it works on and confirms before it starts, and delete asks you to
    type DELETE. Every action is printed while it runs.

.PARAMETER CsvPath
    Use this CSV instead of asking with the Open dialog.

.PARAMETER RootFolder
    Use this redirect root folder instead of asking with the folder dialog.

.PARAMETER WhatIf
    Preview mode. Each operation prints what it WOULD do and changes nothing.

.NOTES
    Run with Windows PowerShell 5.1 (powershell.exe). Windows Forms dialogs need
    an interactive desktop. UNC paths such as \\SERVER\UserRedirection$ work.

.EXAMPLE
    .\namesfolderlookup.ps1

.EXAMPLE
    .\namesfolderlookup.ps1 -WhatIf

.EXAMPLE
    .\namesfolderlookup.ps1 -CsvPath 'C:\csv\names.csv' -RootFolder 'D:\UserRedirection'
#>

param(
    [string]$CsvPath = '',
    [string]$RootFolder = '',
    [switch]$WhatIf
)

$ErrorActionPreference = 'Continue'
$script:Preview = [bool]$WhatIf
$script:CsvFile = ''
$script:RootFolder = ''
$script:Names = @()
$script:Folders = @()
$script:Missing = @()
$script:AdInfo = $null
$script:DestNaming = 'FullName'

$ScriptRoot = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptRoot)) {
    $ScriptRoot = (Get-Location).Path
}
$ProjectRoot = Split-Path -Parent $ScriptRoot
$CsvFolder = Join-Path $ProjectRoot 'csv'
if (-not (Test-Path -LiteralPath $CsvFolder)) {
    $CsvFolder = $ScriptRoot
}

# ============================================
# SMALL HELPERS
# ============================================

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

function Write-Err {
    param([string]$Text)
    Write-Host "    XX  $Text" -ForegroundColor Red
}

function Read-Answer {
    param([string]$Prompt, [string]$Default = '')
    while ($true) {
        if ([string]::IsNullOrWhiteSpace($Default)) {
            $Answer = Read-Host $Prompt
        }
        else {
            $Answer = Read-Host "$Prompt [$Default]"
        }
        if (-not [string]::IsNullOrWhiteSpace($Answer)) { return $Answer.Trim() }
        if (-not [string]::IsNullOrWhiteSpace($Default)) { return $Default }
        Write-Warn 'Please type an answer.'
    }
}

function Read-YesNo {
    param([string]$Prompt)
    while ($true) {
        $Answer = Read-Host "$Prompt [y/N]"
        if ([string]::IsNullOrWhiteSpace($Answer)) { return $false }
        $Text = $Answer.Trim()
        if ($Text -match '^(y|yes)$') { return $true }
        if ($Text -match '^(n|no)$') { return $false }
        Write-Warn 'Answer y (yes) or n (no).'
    }
}

# A folder name uses the username, so names are compared ignoring capitals and
# surrounding spaces. A second, looser key drops dots and spaces as well, which
# catches "Mark Renz D. Bargaso" next to "markrenzdbargaso".
function Get-Key {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return $Text.Trim().ToLowerInvariant()
}

function Get-LooseKey {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return ($Text.Trim().ToLowerInvariant() -replace '[^a-z0-9]', '')
}

function Format-Cell {
    param([string]$Text, [int]$Width)
    if ([string]::IsNullOrEmpty($Text)) { return ''.PadRight($Width) }
    if ($Text.Length -gt $Width) { return ($Text.Substring(0, $Width - 3) + '...') }
    return $Text.PadRight($Width)
}

function Get-StatusColor {
    param([string]$Status)
    switch ($Status) {
        'Matched' { return 'Green' }
        'Check' { return 'Yellow' }
        'No match' { return 'Red' }
        default { return 'White' }
    }
}

function Get-PathKey {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path)) { return '' }
    return ($Path -replace '[\\/]+$', '').ToLowerInvariant()
}

# Write a name the way the CSVs here do: "First Name Middle Initial Last Name".
# A name given as "Last Name, First Name Middle Initial" is turned around, so both
# spellings end up in the same basis.
function Get-DisplayName {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

    $Clean = ($Text -replace '\s+', ' ').Trim()

    if ($Clean -match ',') {
        $Parts = @($Clean -split ',', 2)
        $Last = $Parts[0].Trim()
        $Rest = $Parts[1].Trim()
        $Clean = ('{0} {1}' -f $Rest, $Last).Trim()
    }

    return ($Clean -replace '\s+', ' ').Trim()
}

# A folder name Windows accepts: no invalid characters, no trailing dot or space.
function Get-SafeFolderName {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

    $Name = $Text
    $Bad = [string]([System.IO.Path]::GetInvalidFileNameChars() -join '')
    foreach ($Char in $Bad.ToCharArray()) {
        $Name = $Name.Replace([string]$Char, '')
    }
    $Name = ($Name -replace '\s+', ' ').Trim()
    return $Name.TrimEnd('.', ' ')
}

function Get-DestNamingLabel {
    if ($script:DestNaming -eq 'FullName') {
        return 'full name from the CSV (First Middle Initial Last)'
    }
    return 'username (unchanged)'
}

# The name a folder gets at its destination: the full name from the CSV, or the
# name it already has, depending on the N key in the menu. A folder that matches
# no name keeps its own name.
function Get-DestinationName {
    param($Folder)

    if (-not $Folder) { return '' }
    if ($script:DestNaming -ne 'FullName') { return $Folder.Name }

    $Name = ''
    if ($Folder.Entry) { $Name = Get-DisplayName $Folder.Entry.FullName }
    if ([string]::IsNullOrWhiteSpace($Name)) { $Name = Get-DisplayName $Folder.FullName }

    $Name = Get-SafeFolderName $Name
    if ([string]::IsNullOrWhiteSpace($Name)) { return $Folder.Name }
    return $Name
}

# ============================================
# WINDOWS DIALOGS
# ============================================

function Select-CsvFile {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $Dialog = New-Object System.Windows.Forms.OpenFileDialog
    $Dialog.Title = 'Select the names file (names.csv)'
    $Dialog.Filter = 'CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|All files (*.*)|*.*'
    $Dialog.FilterIndex = 1
    $Dialog.RestoreDirectory = $true
    $Dialog.InitialDirectory = $CsvFolder

    if ($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $Dialog.FileName
    }
    return ''
}

function Select-FolderPath {
    param(
        [string]$Description = 'Select the folder',
        [string]$InitialPath = ''
    )

    Add-Type -AssemblyName System.Windows.Forms

    $Dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $Dialog.Description = $Description
    $Dialog.ShowNewFolderButton = $true

    try {
        # Modern dialog when it can load, classic tree view otherwise
        $Dialog.UseDescriptionForTitle = $true
    }
    catch { }

    if (-not [string]::IsNullOrWhiteSpace($InitialPath) -and (Test-Path -LiteralPath $InitialPath)) {
        try { $Dialog.SelectedPath = $InitialPath } catch { }
    }

    if ($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        if (-not [string]::IsNullOrWhiteSpace($Dialog.SelectedPath)) {
            return $Dialog.SelectedPath
        }
    }
    return ''
}

# ============================================
# READ THE NAMES FILE
# ============================================

function Get-NameListFromCsv {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "The names file was not found: $Path"
    }

    $Rows = @(Import-Csv -LiteralPath $Path -ErrorAction Stop)
    if ($Rows.Count -eq 0) {
        throw "The names file has no rows: $Path"
    }

    $Props = @($Rows[0].PSObject.Properties.Name)

    $UserCol = $Props |
        Where-Object { $_ -match '(?i)^\s*(username|user ?name|user|samaccountname|login|account)\s*$' } |
        Select-Object -First 1

    $FullCol = $Props |
        Where-Object { $_ -match '(?i)^\s*(full ?name|fullname|display ?name|displayname|name)\s*$' } |
        Select-Object -First 1

    $List = New-Object System.Collections.ArrayList
    $NextId = 0
    $NoHeader = $false

    if (-not $UserCol -and -not $FullCol) {
        # No recognized header: treat the single column as one name per row.
        # A file that truly has a "Name" header is covered by the pattern above,
        # so this is a plain list such as "Bargaso, Mark Renz D." on every line.
        $NoHeader = $true
        $OnlyCol = $Props[0]
        if ($OnlyCol -notmatch '(?i)^\s*(names?|full ?name|username|user|samaccountname|display ?name|login)\s*$') {
            $NextId++
            [void]$List.Add([pscustomobject]@{
                Id       = $NextId
                FullName = $OnlyCol
                Username = ''
                AdSource = ''
                AdNote   = ''
            })
        }
        foreach ($Row in $Rows) {
            $Value = [string]$Row.$OnlyCol
            if ([string]::IsNullOrWhiteSpace($Value)) { continue }
            $NextId++
            [void]$List.Add([pscustomobject]@{
                Id       = $NextId
                FullName = $Value.Trim()
                Username = ''
                AdSource = ''
                AdNote   = ''
            })
        }
    }
    else {
        if (-not $UserCol) { $UserCol = '' }
        if (-not $FullCol) { $FullCol = '' }

        foreach ($Row in $Rows) {
            $Full = ''
            $User = ''
            if ($FullCol) { $Full = ([string]$Row.$FullCol).Trim() }
            if ($UserCol) { $User = ([string]$Row.$UserCol).Trim() }
            if ([string]::IsNullOrWhiteSpace($Full) -and [string]::IsNullOrWhiteSpace($User)) { continue }
            $NextId++
            [void]$List.Add([pscustomobject]@{
                Id       = $NextId
                FullName = $Full
                Username = $User
                AdSource = ''
                AdNote   = ''
            })
        }
    }

    if ($List.Count -eq 0) {
        throw "No names could be read from $Path"
    }

    $Script:CsvColumns = @{
        Username = $UserCol
        FullName = $FullCol
        NoHeader = $NoHeader
    }

    return $List.ToArray()
}

# ============================================
# ASK AD FOR THE REAL USERNAME OF EACH NAME
# ============================================

function Set-NameField {
    param($Object, [string]$Field, $Value)
    if ($null -eq $Object) { return }
    if ($Object.PSObject.Properties[$Field]) {
        $Object.$Field = $Value
    }
    else {
        $Object | Add-Member -NotePropertyName $Field -NotePropertyValue $Value -Force
    }
}

# The username this project's AD accounts follow: initials + surname.
#   "Michael Angelo C. Jordan"  -> "macjordan"
#   "Jordan, Michael Angelo C." -> "macjordan"
function Get-UsernameCandidate {
    param([string]$FullName)

    if ([string]::IsNullOrWhiteSpace($FullName)) { return '' }

    $Text = $FullName.Trim() -replace '\s+', ' '
    $Surname = ''
    $Given = @()

    if ($Text -match ',') {
        $Parts = @($Text -split ',', 2)
        $Surname = $Parts[0].Trim()
        $Given = @($Parts[1].Trim() -split ' ')
    }
    else {
        $Words = @($Text -split ' ' | Where-Object { $_ })
        if ($Words.Count -ge 2) {
            $Surname = $Words[-1]
            $Given = @($Words[0..($Words.Count - 2)])
        }
        else {
            $Surname = $Text
        }
    }

    $Initials = ''
    foreach ($Word in $Given) {
        $Letters = ($Word -replace '[^A-Za-z]', '')
        if ($Letters.Length -eq 0) { continue }
        $Initials += $Letters.Substring(0, 1)
    }

    $Built = $Initials + ($Surname -replace '[^A-Za-z]', '')
    return ($Built.ToLowerInvariant() -replace '[^a-z0-9]', '')
}

# Find the AD account behind a full name. Returns the accounts found (may be empty).
function Find-ADUserByFullName {
    param([string]$FullName)

    if ([string]::IsNullOrWhiteSpace($FullName)) { return @() }

    $Props = @('DisplayName', 'Name', 'SamAccountName', 'Enabled')
    $Escaped = $FullName.Trim().Replace("'", "''")

    # 1. the name exactly as written in the CSV
    foreach ($Filter in @("DisplayName -eq '$Escaped'", "Name -eq '$Escaped'")) {
        try { $Hit = @(Get-ADUser -Filter $Filter -Properties $Props -ErrorAction Stop) }
        catch { $Hit = @() }
        if ($Hit.Count -gt 0) { return $Hit }
    }

    # 2. the same person written the other way round: "Surname, First" <-> "First Surname"
    $Swapped = ''
    if ($FullName -match ',') {
        $Parts = @($FullName -split ',', 2)
        $Swapped = ('{0} {1}' -f ($Parts[1].Trim()), ($Parts[0].Trim()))
    }
    else {
        $Words = @($FullName.Trim() -split '\s+' | Where-Object { $_ })
        if ($Words.Count -ge 2) {
            $Swapped = ('{0}, {1}' -f $Words[-1], ($Words[0..($Words.Count - 2)] -join ' ').Trim())
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($Swapped)) {
        $SwappedClean = ($Swapped -replace '\s+', ' ').Trim()
        $EscapedSwap = $SwappedClean.Replace("'", "''")
        foreach ($Filter in @("DisplayName -eq '$EscapedSwap'", "Name -eq '$EscapedSwap'")) {
            try { $Hit = @(Get-ADUser -Filter $Filter -Properties $Props -ErrorAction Stop) }
            catch { $Hit = @() }
            if ($Hit.Count -gt 0) { return $Hit }
        }
    }

    # 3. exactly one account whose name contains the whole name
    try { $Like = @(Get-ADUser -Filter "Name -like '*$Escaped*'" -Properties $Props -ErrorAction Stop) }
    catch { $Like = @() }
    if ($Like.Count -eq 1) { return $Like }

    return @()
}

# Fill in the real username of every name in the CSV: from the CSV itself, from
# Active Directory, or - when AD has no such account - built from the full name.
function Resolve-NameUsernames {
    param([switch]$Quiet)

    $HaveAd = [bool](Get-Module -Name ActiveDirectory -ListAvailable)
    if ($HaveAd) {
        try { Import-Module ActiveDirectory -ErrorAction Stop }
        catch { $HaveAd = $false }
    }

    if ($HaveAd) {
        $script:AdInfo = 'ok'
        if (-not $Quiet) { Write-Step 'Looking each name up in Active Directory (this can take a moment)' }
    }
    else {
        $script:AdInfo = 'unavailable'
        if (-not $Quiet) {
            Write-Warn 'Active Directory cannot be read on this computer.'
            Write-Info 'AD accounts are skipped - usernames are built from the full names instead.'
        }
    }

    $FromCsv = 0
    $FromAd = 0
    $Guessed = 0
    $Index = 0

    foreach ($Name in $script:Names) {
        $Index++
        Set-NameField -Object $Name -Field 'AdNote' -Value ''

        # The CSV already names the user - that wins.
        if (-not [string]::IsNullOrWhiteSpace($Name.Username)) {
            Set-NameField -Object $Name -Field 'AdSource' -Value 'CSV'
            $FromCsv++
            continue
        }

        $Found = @()
        if ($HaveAd -and -not [string]::IsNullOrWhiteSpace($Name.FullName)) {
            $Found = @(Find-ADUserByFullName -FullName $Name.FullName)
        }

        if ($Found.Count -gt 0) {
            $Enabled = @($Found | Where-Object { $_.Enabled })
            $Pick = if ($Enabled.Count -gt 0) { $Enabled[0] } else { $Found[0] }
            Set-NameField -Object $Name -Field 'Username' -Value ([string]$Pick.SamAccountName)
            Set-NameField -Object $Name -Field 'AdSource' -Value 'AD'
            if ($Found.Count -gt 1) {
                Set-NameField -Object $Name -Field 'AdNote' -Value ("$($Found.Count) AD accounts share this name - using $($Pick.SamAccountName)")
            }
            $FromAd++
        }
        else {
            # No AD account for this name - build the username from the name itself
            $Built = Get-UsernameCandidate -FullName $Name.FullName
            Set-NameField -Object $Name -Field 'Username' -Value $Built
            Set-NameField -Object $Name -Field 'AdSource' -Value 'Guessed'
            if ($HaveAd) {
                Set-NameField -Object $Name -Field 'AdNote' -Value 'no AD account found for this name'
            }
            elseif (-not $Built) {
                Set-NameField -Object $Name -Field 'AdNote' -Value 'the name is too short to build a username'
            }
            $Guessed++
        }

        if ($HaveAd -and -not $Quiet -and $script:Names.Count -ge 40 -and ($Index % 25) -eq 0) {
            Write-Info "checked $Index of $($script:Names.Count) names..."
        }
    }

    if ($HaveAd -and -not $Quiet) {
        Write-Ok "Usernames ready: $FromCsv from the CSV, $FromAd found in AD, $Guessed built from the name."
    }

    return $true
}

# ============================================
# SCAN THE REDIRECT ROOT AND MATCH EVERY FOLDER
# ============================================

# How a folder was matched, for the report
function Get-HowLabel {
    param($MatchName)
    if (-not $MatchName) { return '' }
    switch ($MatchName.AdSource) {
        'AD' { return 'AD username' }
        'CSV' { return 'CSV username' }
        default { return 'built from the name' }
    }
}

function Update-Report {

    $script:Folders = @()
    $script:Missing = @()

    if (-not (Test-Path -LiteralPath $script:RootFolder)) {
        Write-Err "The redirect folder cannot be opened: $($script:RootFolder)"
        return $false
    }

    # The folders in a redirection root are named after the AD username, so the
    # username resolved for each name in the CSV is what a folder is matched on.
    $ExactMap = @{}
    $LooseMap = @{}
    $UserMap = @{}
    foreach ($Name in $script:Names) {
        $UserKey = Get-Key $Name.Username
        if ($UserKey.Length -gt 0 -and -not $UserMap.ContainsKey($UserKey)) { $UserMap[$UserKey] = $Name }
        foreach ($Text in @($Name.FullName, $Name.Username)) {
            $Key = Get-Key $Text
            if ($Key.Length -eq 0) { continue }
            if (-not $ExactMap.ContainsKey($Key)) { $ExactMap[$Key] = $Name }
            $Loose = Get-LooseKey $Text
            if ($Loose.Length -gt 0 -and -not $LooseMap.ContainsKey($Loose)) { $LooseMap[$Loose] = $Name }
        }
    }

    Write-Step "Reading the folders inside $($script:RootFolder)"
    try {
        $Dirs = @(Get-ChildItem -LiteralPath $script:RootFolder -Directory -ErrorAction Stop |
                  Sort-Object -Property Name)
    }
    catch {
        Write-Err "Could not list the folders: $($_.Exception.Message)"
        return $false
    }

    if ($Dirs.Count -eq 0) {
        Write-Warn 'There are no folders directly inside that folder.'
    }
    else {
        Write-Ok "$($Dirs.Count) folder(s) found."
    }

    $Rows = New-Object System.Collections.ArrayList
    $MatchedIds = New-Object System.Collections.ArrayList
    $Number = 0

    foreach ($Dir in $Dirs) {
        $Entry = $null
        $Status = 'No match'
        $How = ''
        $Key = Get-Key $Dir.Name

        # 1. the folder name is the AD username of a name in the CSV
        if ($UserMap.ContainsKey($Key)) {
            $Entry = $UserMap[$Key]
            if ($Entry.AdSource -eq 'AD') { $Status = 'Matched' } else { $Status = 'Check' }
            $How = Get-HowLabel -MatchName $Entry
        }
        # 2. the folder name is literally a name from the CSV
        elseif ($ExactMap.ContainsKey($Key)) {
            $Entry = $ExactMap[$Key]
            $Status = 'Matched'
            $How = 'full name'
        }
        else {
            $LooseKey = Get-LooseKey $Dir.Name
            if ($LooseKey.Length -gt 0 -and $LooseMap.ContainsKey($LooseKey)) {
                $Entry = $LooseMap[$LooseKey]
                $Status = 'Check'
                $How = 'loose name'
            }
            elseif ($LooseKey.Length -ge 6) {
                # Only one name could be inside this folder name - show it as a hint
                $Candidates = @($LooseMap.Keys | Where-Object { $_.Length -ge 6 -and $LooseKey -like "*$_*" })
                if ($Candidates.Count -eq 1) {
                    $Entry = $LooseMap[$Candidates[0]]
                    $Status = 'Check'
                    $How = 'part of the name'
                }
            }
        }

        if ($Entry) { [void]$MatchedIds.Add($Entry.Id) }

        $Number++
        [void]$Rows.Add([pscustomobject]@{
            Number   = $Number
            Name     = $Dir.Name
            Path     = $Dir.FullName
            Status   = $Status
            How      = $How
            FullName = if ($Entry) { $Entry.FullName } else { '' }
            Username = if ($Entry) { $Entry.Username } else { '' }
            Entry    = $Entry
        })
    }

    $script:Folders = $Rows.ToArray()

    # Names from the CSV that have no folder at all
    $MissingList = New-Object System.Collections.ArrayList
    foreach ($Name in $script:Names) {
        if (-not ($MatchedIds -contains $Name.Id)) {
            [void]$MissingList.Add([pscustomobject]@{
                FullName = $Name.FullName
                Username = $Name.Username
                AdSource = $Name.AdSource
                AdNote   = $Name.AdNote
            })
        }
    }
    $script:Missing = $MissingList.ToArray()

    return $true
}

function Show-Report {

    $Matched = @($script:Folders | Where-Object { $_.Status -eq 'Matched' })
    $Check = @($script:Folders | Where-Object { $_.Status -eq 'Check' })
    $NoMatch = @($script:Folders | Where-Object { $_.Status -eq 'No match' })

    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host ' MATCH REPORT' -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host ("  Names file     : {0}" -f $script:CsvFile) -ForegroundColor White
    Write-Host ("  Redirect folder: {0}" -f $script:RootFolder) -ForegroundColor White
    Write-Host ("  Names in CSV   : {0}    Folders found: {1}" -f $script:Names.Count, $script:Folders.Count) -ForegroundColor White
    Write-Host ("  Matched {0}   Check {1}   No match {2}   Name without folder {3}" -f $Matched.Count, $Check.Count, $NoMatch.Count, $script:Missing.Count) -ForegroundColor Gray

    if ($script:Folders.Count -gt 0) {
        Write-Host ''
        Write-Host ' FOLDERS' -ForegroundColor Cyan
        Write-Host ('  {0,3})  {1}  {2}  {3}' -f '#', (Format-Cell 'FOLDER NAME' 34), (Format-Cell 'STATUS' 9), 'FULL NAME / USERNAME') -ForegroundColor DarkCyan
        foreach ($Folder in $script:Folders) {
            $Who = $Folder.FullName
            if ($Folder.Username) {
                $Who = ('{0} [{1}]' -f $Who, $Folder.Username).Trim()
            }
            $Line = '  {0,3})  {1}  {2}  {3}' -f $Folder.Number,
                (Format-Cell $Folder.Name 34),
                (Format-Cell $Folder.Status 9),
                $Who
            Write-Host $Line -ForegroundColor (Get-StatusColor -Status $Folder.Status)
            if ($Folder.How) {
                Write-Host ('         matched on {0}' -f $Folder.How) -ForegroundColor DarkGray
            }
        }
    }

    if ($script:Missing.Count -gt 0) {
        Write-Host ''
        Write-Host ' NAMES FROM THE CSV THAT HAVE NO FOLDER (nothing to copy or delete)' -ForegroundColor Cyan
        foreach ($Name in $script:Missing) {
            $Label = if ($Name.FullName) { $Name.FullName } else { $Name.Username }
            $Extra = ''
            if ($Name.Username) { $Extra = " -> folder would be '$($Name.Username)'" }
            if ($Name.AdNote) { $Extra = "$Extra  ($($Name.AdNote))" }
            Write-Host ("   -  {0}{1}" -f $Label, $Extra) -ForegroundColor Yellow
        }
    }

    if ($script:Folders.Count -gt 0 -and $Matched.Count -eq 0 -and $Check.Count -eq 0) {
        Write-Host ''
        Write-Warn 'No folder matched a single name in the CSV. What to check:'
        Write-Info '1. Is this really the root folder? It must be the one that holds the per-user folders.'
        Write-Info '2. Compare the resolved usernames with the folder names shown above.'
        Write-Info '3. If AD could not be read, the usernames were only built from the names.'
    }

    Write-Host ''
    Write-Host '  Matched  = the folder name is the AD username of a name in the CSV' -ForegroundColor Gray
    Write-Host '  Check    = the names look alike or the username was built from the name - look at it first' -ForegroundColor Gray
    Write-Host '  No match = no name in the CSV fits this folder (leftover folder)' -ForegroundColor Gray
}

# ============================================
# PICK THE FOLDERS AN OPERATION WORKS ON
# ============================================

function Select-TargetFolders {
    param([string]$Action)

    if ($script:Folders.Count -eq 0) {
        Write-Warn 'There are no folders to work on.'
        return @()
    }

    $Matched = @($script:Folders | Where-Object { $_.Status -eq 'Matched' -or $_.Status -eq 'Check' })
    $NoMatch = @($script:Folders | Where-Object { $_.Status -eq 'No match' })

    Write-Host ''
    Write-Host (" FOLDERS FOR: {0}" -f $Action.ToUpperInvariant()) -ForegroundColor Cyan
    foreach ($Folder in $script:Folders) {
        $Who = $Folder.FullName
        if ($Folder.Username) { $Who = ('{0} [{1}]' -f $Who, $Folder.Username).Trim() }
        Write-Host ('  {0,3})  {1}  {2}  {3}' -f $Folder.Number,
            (Format-Cell $Folder.Name 34),
            (Format-Cell $Folder.Status 9),
            $Who) -ForegroundColor (Get-StatusColor -Status $Folder.Status)
    }

    Write-Host ''
    Write-Host '  A) all folders that match a name in the CSV (Matched + Check)' -ForegroundColor Cyan
    Write-Host ("  U) the folders that match NO name in the CSV ({0} folder(s))" -f $NoMatch.Count) -ForegroundColor Cyan
    Write-Host '  Enter = cancel' -ForegroundColor Cyan

    while ($true) {
        $Answer = Read-Host "`nType numbers (1,3,5), A, U or Enter to cancel"
        if ([string]::IsNullOrWhiteSpace($Answer)) { return @() }

        $Text = $Answer.Trim()

        if ($Text -match '^(a|all)$') {
            if ($Matched.Count -eq 0) {
                Write-Warn 'No folder matches a name in the CSV.'
                continue
            }
            return @($Matched)
        }

        if ($Text -match '^(u|unmatched)$') {
            if ($NoMatch.Count -eq 0) {
                Write-Warn 'Every folder matches a name in the CSV - nothing to pick with U.'
                continue
            }
            Write-Warn "U picks the $($NoMatch.Count) folder(s) that match NO name."
            return @($NoMatch)
        }

        $Parts = @($Text -split '[,; ]+' | Where-Object { $_ -ne '' })
        $Selected = New-Object System.Collections.ArrayList
        $Valid = $true

        foreach ($Part in $Parts) {
            $Index = 0
            if (-not [int]::TryParse($Part, [ref]$Index)) {
                Write-Warn "'$Part' is not a number."
                $Valid = $false
                break
            }
            if ($Index -lt 1 -or $Index -gt $script:Folders.Count) {
                Write-Warn "There is no folder number $Index."
                $Valid = $false
                break
            }
            [void]$Selected.Add($script:Folders[$Index - 1])
        }

        if (-not $Valid) { continue }
        return @($Selected | Sort-Object -Property Number)
    }
}

function Confirm-Operation {
    param([string]$Question)
    if ($script:Preview) {
        Write-Info 'Preview mode is on: nothing will be changed.'
        return $true
    }
    return (Read-YesNo -Prompt $Question)
}

# ============================================
# OPERATIONS
# ============================================

function Invoke-CopyFolders {

    $Selected = @(Select-TargetFolders -Action 'copy')
    if ($Selected.Count -eq 0) { Write-Info 'Nothing selected.'; return }

    $Destination = Select-FolderPath -Description 'Where should the copies be placed?'
    if ([string]::IsNullOrWhiteSpace($Destination)) {
        Write-Warn 'No destination picked - cancelled.'
        return
    }

    Write-Host ''
    Write-Host (" {0} folder(s) will be copied to:" -f $Selected.Count) -ForegroundColor White
    Write-Info $Destination

    $Overwrite = Read-YesNo -Prompt 'Replace a folder that already exists at the destination?'
    if (-not (Confirm-Operation -Question 'Start the copy')) { Write-Info 'Cancelled.'; return }

    $Copied = 0
    $Skipped = 0
    $Failed = 0

    foreach ($Folder in $Selected) {
        $NewName = Get-DestinationName -Folder $Folder
        $Target = Join-Path -Path $Destination -ChildPath $NewName

        if (Test-Path -LiteralPath $Target) {
            if (-not $Overwrite) {
                Write-Warn "already there, skipped: $Target"
                $Skipped++
                continue
            }
            if (-not $script:Preview) {
                try { Remove-Item -LiteralPath $Target -Recurse -Force -ErrorAction Stop }
                catch {
                    Write-Err "could not replace the old folder: $Target - $($_.Exception.Message)"
                    $Failed++
                    continue
                }
            }
        }

        if ($script:Preview) {
            Write-Host ("  WOULD copy {0}  ->  {1}" -f $Folder.Path, $Target) -ForegroundColor Yellow
            $Copied++
            continue
        }

        # Copy into a throw-away name first, then rename it, so the copy lands
        # under the full name even though the original folder is named after the
        # AD username.
        $Staging = ''
        try {
            $Staging = Join-Path -Path $Destination -ChildPath ('~nf-copy-' + [guid]::NewGuid().ToString('N'))
            Copy-Item -LiteralPath $Folder.Path -Destination $Staging -Recurse -Force -ErrorAction Stop
            Rename-Item -LiteralPath $Staging -NewName $NewName -ErrorAction Stop
            Write-Host ("  copied  {0}  ->  {1}" -f $Folder.Name, $NewName) -ForegroundColor Green
            $Copied++
        }
        catch {
            Write-Err "copy failed: $($Folder.Name) - $($_.Exception.Message)"
            if ($Staging -and (Test-Path -LiteralPath $Staging)) {
                try { Remove-Item -LiteralPath $Staging -Recurse -Force -ErrorAction Stop } catch { }
            }
            $Failed++
        }
    }

    Write-Host ''
    Write-Host (" Copy result: {0} copied, {1} skipped, {2} failed" -f $Copied, $Skipped, $Failed) -ForegroundColor Cyan
}

function Invoke-MoveFolders {

    $Selected = @(Select-TargetFolders -Action 'move')
    if ($Selected.Count -eq 0) { Write-Info 'Nothing selected.'; return }

    $Destination = Select-FolderPath -Description 'Where should the folders be moved to?'
    if ([string]::IsNullOrWhiteSpace($Destination)) {
        Write-Warn 'No destination picked - cancelled.'
        return
    }

    $DestinationKey = Get-PathKey $Destination
    $RootKey = Get-PathKey $script:RootFolder
    if ($DestinationKey -eq $RootKey) {
        Write-Warn 'That is the same folder they are already in - cancelled.'
        return
    }
    if ($DestinationKey -like "$RootKey\*") {
        Write-Warn 'That folder is INSIDE the redirect root. Moving folders there would hide them from this report - cancelled.'
        return
    }

    Write-Host ''
    Write-Host (" {0} folder(s) will be MOVED to:" -f $Selected.Count) -ForegroundColor White
    Write-Info $Destination
    Write-Host ' Moving takes the folder away from the redirect root.' -ForegroundColor Yellow

    if (-not (Confirm-Operation -Question 'Move these folders')) { Write-Info 'Cancelled.'; return }

    $Moved = 0
    $Skipped = 0
    $Failed = 0

    foreach ($Folder in $Selected) {
        $NewName = Get-DestinationName -Folder $Folder
        $Target = Join-Path -Path $Destination -ChildPath $NewName

        if (Test-Path -LiteralPath $Target) {
            Write-Warn "already there, skipped: $Target"
            $Skipped++
            continue
        }

        if ($script:Preview) {
            Write-Host ("  WOULD move {0}  ->  {1}" -f $Folder.Path, $Target) -ForegroundColor Yellow
            $Moved++
            continue
        }

        try {
            Move-Item -LiteralPath $Folder.Path -Destination $Target -Force
            Write-Host ("  moved   {0}  ->  {1}" -f $Folder.Name, $NewName) -ForegroundColor Green
            $Moved++
        }
        catch {
            Write-Err "move failed: $($Folder.Name) - $($_.Exception.Message)"
            $Failed++
        }
    }

    Write-Host ''
    Write-Host (" Move result: {0} moved, {1} skipped, {2} failed" -f $Moved, $Skipped, $Failed) -ForegroundColor Cyan
    if (-not $script:Preview -and $Moved -gt 0) {
        Write-Info 'Run R (rescan) so the report shows the new state.'
    }
}

function Invoke-RenameFolders {

    $Selected = @(Select-TargetFolders -Action 'rename to the full name from the CSV (First Middle Initial Last)')
    if ($Selected.Count -eq 0) { Write-Info 'Nothing selected.'; return }

    $Plan = New-Object System.Collections.ArrayList
    foreach ($Folder in $Selected) {
        if (-not $Folder.Entry) {
            Write-Warn "no name in the CSV for this folder, skipped: $($Folder.Name)"
            continue
        }
        $NewName = Get-DisplayName $Folder.Entry.FullName
        if ([string]::IsNullOrWhiteSpace($NewName)) { $NewName = Get-DisplayName $Folder.Entry.Username }
        $NewName = Get-SafeFolderName $NewName
        if ([string]::IsNullOrWhiteSpace($NewName)) {
            Write-Warn "no usable name for this folder, skipped: $($Folder.Name)"
            continue
        }
        if ($NewName -eq $Folder.Name) {
            Write-Info "already named that way, skipped: $($Folder.Name)"
            continue
        }
        $Target = Join-Path -Path (Split-Path -Parent $Folder.Path) -ChildPath $NewName
        if (Test-Path -LiteralPath $Target) {
            Write-Warn "a folder with the new name already exists, skipped: $NewName"
            continue
        }
        [void]$Plan.Add([pscustomobject]@{
            OldPath = $Folder.Path
            OldName = $Folder.Name
            NewName = $NewName
        })
    }

    if ($Plan.Count -eq 0) { Write-Warn 'Nothing to rename.'; return }

    Write-Host ''
    Write-Host " Rename plan ($($Plan.Count) folder(s)):" -ForegroundColor White
    foreach ($Item in $Plan) {
        Write-Info ("{0}  ->  {1}" -f $Item.OldName, $Item.NewName)
    }

    if (-not (Confirm-Operation -Question 'Rename these folders')) { Write-Info 'Cancelled.'; return }

    $Renamed = 0
    $Failed = 0
    foreach ($Item in $Plan) {
        if ($script:Preview) {
            Write-Host ("  WOULD rename {0}  ->  {1}" -f $Item.OldName, $Item.NewName) -ForegroundColor Yellow
            $Renamed++
            continue
        }
        try {
            Rename-Item -LiteralPath $Item.OldPath -NewName $Item.NewName -ErrorAction Stop
            Write-Host ("  renamed {0}  ->  {1}" -f $Item.OldName, $Item.NewName) -ForegroundColor Green
            $Renamed++
        }
        catch {
            Write-Err "rename failed: $($Item.OldName) - $($_.Exception.Message)"
            $Failed++
        }
    }

    Write-Host ''
    Write-Host (" Rename result: {0} renamed, {1} failed" -f $Renamed, $Failed) -ForegroundColor Cyan
}

function Invoke-DeleteFolders {

    $Selected = @(Select-TargetFolders -Action 'delete')
    if ($Selected.Count -eq 0) { Write-Info 'Nothing selected.'; return }

    Write-Host ''
    Write-Host ' THESE FOLDERS WILL BE DELETED WITH EVERYTHING INSIDE THEM:' -ForegroundColor Red
    $TotalFiles = 0
    foreach ($Folder in $Selected) {
        $FileCount = 0
        try {
            $FileCount = @(Get-ChildItem -LiteralPath $Folder.Path -File -Recurse -Force -ErrorAction SilentlyContinue).Count
        }
        catch { }
        $TotalFiles += $FileCount
        Write-Host ("   {0}  ({1} file(s))" -f $Folder.Name, $FileCount) -ForegroundColor Red
    }
    Write-Host ''
    Write-Host (" {0} folder(s), about {1} file(s) in total." -f $Selected.Count, $TotalFiles) -ForegroundColor Yellow

    if (-not $script:Preview) {
        $Confirm = Read-Host "`nType DELETE to confirm (anything else cancels)"
        if ($Confirm -ne 'DELETE') {
            Write-Info 'Cancelled. Nothing was deleted.'
            return
        }
    }
    else {
        Write-Info 'Preview mode is on: nothing will be deleted.'
    }

    $RootKey = Get-PathKey $script:RootFolder
    $Deleted = 0
    $Failed = 0

    foreach ($Folder in $Selected) {
        $Key = Get-PathKey $Folder.Path
        if ([string]::IsNullOrEmpty($Key) -or $Key -eq $RootKey) {
            Write-Warn "refused - this is the redirect root itself: $($Folder.Path)"
            $Failed++
            continue
        }

        if ($script:Preview) {
            Write-Host ("  WOULD delete {0}" -f $Folder.Path) -ForegroundColor Yellow
            $Deleted++
            continue
        }

        try {
            Remove-Item -LiteralPath $Folder.Path -Recurse -Force -ErrorAction Stop
            Write-Host ("  deleted {0}" -f $Folder.Name) -ForegroundColor Green
            $Deleted++
        }
        catch {
            Write-Err "delete failed: $($Folder.Name) - $($_.Exception.Message)"
            $Failed++
        }
    }

    Write-Host ''
    Write-Host (" Delete result: {0} deleted, {1} failed" -f $Deleted, $Failed) -ForegroundColor Cyan
    if (-not $script:Preview -and $Deleted -gt 0) {
        Write-Info 'Run R (rescan) so the report shows the new state.'
    }
}

function Invoke-ExportReport {

    Add-Type -AssemblyName System.Windows.Forms

    $Dialog = New-Object System.Windows.Forms.SaveFileDialog
    $Dialog.Title = 'Where should the report be saved?'
    $Dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $Dialog.FileName = 'Folder_Match_Report_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    $Dialog.InitialDirectory = Split-Path -Parent $script:CsvFile

    if ($Dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Warn 'No place picked - cancelled.'
        return
    }

    $Rows = New-Object System.Collections.ArrayList
    foreach ($Folder in $script:Folders) {
        [void]$Rows.Add([pscustomobject]@{
            Status   = $Folder.Status
            Folder   = $Folder.Name
            NewFolderName = Get-DestinationName -Folder $Folder
            FullName = $Folder.FullName
            Username = $Folder.Username
            MatchedBy = $Folder.How
            Path     = $Folder.Path
        })
    }
    foreach ($Name in $script:Missing) {
        [void]$Rows.Add([pscustomobject]@{
            Status   = 'Name without folder'
            Folder   = ''
            FullName = $Name.FullName
            Username = $Name.Username
            MatchedBy = $Name.AdNote
            Path     = ''
        })
    }

    try {
        $Rows | Export-Csv -LiteralPath $Dialog.FileName -NoTypeInformation -Encoding UTF8
        Write-Ok "Report saved to $($Dialog.FileName)"
        Write-Info 'Open it in Excel to filter the No match and Name without folder rows.'
    }
    catch {
        Write-Err "Could not save the report: $($_.Exception.Message)"
    }
}

# ============================================
# HEADER AND MENU
# ============================================

function Show-Header {
    Clear-Host
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host ' NAMES.CSV  <->  FOLDER REDIRECTION LOOKUP' -ForegroundColor Yellow
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host ("  Names file     : {0}" -f $script:CsvFile) -ForegroundColor Gray
    Write-Host ("  Redirect folder: {0}" -f $script:RootFolder) -ForegroundColor Gray
    Write-Host ("  New folder name: {0}" -f (Get-DestNamingLabel)) -ForegroundColor Gray
    if ($script:Preview) {
        Write-Host '  Mode           : PREVIEW (-WhatIf) - nothing is changed' -ForegroundColor Yellow
    }
    else {
        Write-Host '  Mode           : LIVE - every operation asks before it starts' -ForegroundColor Green
    }
}

function Show-Menu {
    Write-Host ''
    Write-Host ('-' * 78) -ForegroundColor Cyan
    Write-Host '  1) Copy folders to another folder' -ForegroundColor White
    Write-Host '  2) Move folders to another folder' -ForegroundColor White
    Write-Host '  3) Rename folders to the full name from the CSV' -ForegroundColor White
    Write-Host '  4) Delete folders (asks you to type DELETE)' -ForegroundColor Red
    Write-Host '  5) Export the report to CSV' -ForegroundColor White
    Write-Host ("  N) New folder name at the destination: {0}" -f (Get-DestNamingLabel)) -ForegroundColor Cyan
    Write-Host '  R) Rescan the redirect folder' -ForegroundColor Cyan
    Write-Host '  S) Choose another CSV or another redirect folder' -ForegroundColor Cyan
    Write-Host '  Q) Exit' -ForegroundColor Cyan
}

# ============================================
# START: ASK FOR THE TWO INPUTS
# ============================================

Show-Header

if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    Write-Step 'Pick the names file (names.csv)'
    $CsvPath = Select-CsvFile
    if ([string]::IsNullOrWhiteSpace($CsvPath)) {
        Write-Warn 'No names file selected. Nothing to do.'
        Read-Host 'Press Enter to close'
        exit 0
    }
}
if (-not (Test-Path -LiteralPath $CsvPath)) {
    Write-Err "The names file was not found: $CsvPath"
    Read-Host 'Press Enter to close'
    exit 1
}
$script:CsvFile = (Resolve-Path -LiteralPath $CsvPath).Path
Write-Ok "Names file: $($script:CsvFile)"

if ([string]::IsNullOrWhiteSpace($RootFolder)) {
    Write-Step 'Pick the main folder redirection folder (the root that holds the user folders)'
    $RootFolder = Select-FolderPath -Description 'Select the folder redirection root folder'
    if ([string]::IsNullOrWhiteSpace($RootFolder)) {
        Write-Warn 'No redirect folder selected. Nothing to do.'
        Read-Host 'Press Enter to close'
        exit 0
    }
}
if (-not (Test-Path -LiteralPath $RootFolder)) {
    Write-Err "That redirect folder cannot be opened: $RootFolder"
    Read-Host 'Press Enter to close'
    exit 1
}
$script:RootFolder = $RootFolder
Write-Ok "Redirect folder: $($script:RootFolder)"

Write-Step "Reading the names in $($script:CsvFile)"
try {
    $script:Names = Get-NameListFromCsv -Path $script:CsvFile
}
catch {
    Write-Err $_.Exception.Message
    Read-Host 'Press Enter to close'
    exit 1
}
Write-Ok "$($script:Names.Count) name(s) read."
if ($script:CsvColumns.NoHeader) {
    Write-Info 'The CSV has no Username / FullName header, so every line is used as a name.'
}

if (-not (Resolve-NameUsernames)) {
    Read-Host 'Press Enter to close'
    exit 1
}

if (-not (Update-Report)) {
    Read-Host 'Press Enter to close'
    exit 1
}

# ============================================
# MAIN LOOP
# ============================================

while ($true) {

    Show-Header
    Show-Report
    Show-Menu

    $Choice = Read-Host "`nChoose a number, N, R, S or Q"
    if ([string]::IsNullOrWhiteSpace($Choice)) { continue }
    $Text = $Choice.Trim()

    switch -Regex ($Text) {

        '^(q|quit|exit)$' {
            break
        }

        '^(r|rescan)$' {
            [void](Update-Report)
            continue
        }

        '^(s|source|sources)$' {
            Write-Host ''
            $Sub = Read-Host 'C = another names file, F = another redirect folder, B = both, Enter = nothing'
            if ([string]::IsNullOrWhiteSpace($Sub)) { continue }
            $SubText = $Sub.Trim()

            if ($SubText -match '^(c|f|b)$') {
                if ($SubText -match '^(c|b)$') {
                    $NewCsv = Select-CsvFile
                    if ([string]::IsNullOrWhiteSpace($NewCsv)) {
                        Write-Warn 'No names file selected - the old one is kept.'
                    }
                    else {
                        try {
                            $Names = Get-NameListFromCsv -Path $NewCsv
                            $script:CsvFile = (Resolve-Path -LiteralPath $NewCsv).Path
                            $script:Names = $Names
                            Write-Ok "Names file: $($script:CsvFile) ($($script:Names.Count) name(s))"
                            [void](Resolve-NameUsernames)
                        }
                        catch {
                            Write-Err $_.Exception.Message
                        }
                    }
                }
                if ($SubText -match '^(f|b)$') {
                    $NewRoot = Select-FolderPath -Description 'Select the folder redirection root folder'
                    if ([string]::IsNullOrWhiteSpace($NewRoot)) {
                        Write-Warn 'No redirect folder selected - the old one is kept.'
                    }
                    elseif (-not (Test-Path -LiteralPath $NewRoot)) {
                        Write-Err "That folder cannot be opened: $NewRoot"
                    }
                    else {
                        $script:RootFolder = $NewRoot
                        Write-Ok "Redirect folder: $($script:RootFolder)"
                    }
                }
                [void](Update-Report)
            }
            else {
                Write-Warn 'Type C, F, B or press Enter.'
            }
            continue
        }

        '^1$' {
            Invoke-CopyFolders
            Read-Host 'Press Enter to go back to the menu'
            continue
        }

        '^2$' {
            Invoke-MoveFolders
            Read-Host 'Press Enter to go back to the menu'
            continue
        }

        '^3$' {
            Invoke-RenameFolders
            Read-Host 'Press Enter to go back to the menu'
            continue
        }

        '^4$' {
            Invoke-DeleteFolders
            Read-Host 'Press Enter to go back to the menu'
            continue
        }

        '^5$' {
            Invoke-ExportReport
            Read-Host 'Press Enter to go back to the menu'
            continue
        }

        '^(n|name)$' {
            if ($script:DestNaming -eq 'FullName') { $script:DestNaming = 'Username' }
            else { $script:DestNaming = 'FullName' }
            continue
        }

        default {
            Write-Warn "'$Text' is not one of the choices. Try 1-5, N, R, S or Q."
            Start-Sleep -Milliseconds 1200
            continue
        }
    }

    break
}

Clear-Host
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host ' CLOSED - nothing was left running' -ForegroundColor Green
Write-Host ('=' * 78) -ForegroundColor Cyan
