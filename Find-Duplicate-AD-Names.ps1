#Requires -Version 5.1
<#
.SYNOPSIS
    Read-only scan of all on-premises Active Directory users for duplicate or
    near-duplicate display names. Changes nothing in AD.

.DESCRIPTION
    Loads every user from the domain (or one OU), then reports:
      * Exact     - two or more users whose names are the same after tidying
                    (extra spaces, capital letters and punctuation are ignored).
      * Reordered - the same words in a different order, such as
                    "Goingco, Ejay" and "Ejay Goingco" (word order ignored).
      * Near      - names that look alike, such as "Jon Smith" and "John Smith",
                    or "Jane Smith" and "Janet Smith".

    The result is shown on screen and saved as a CSV file you can open in Excel.
    The CSV is saved in the "Output Report" folder next to this script.

    Only read access is needed (Get-ADUser / Get-ADOrganizationalUnit), and a
    normal domain user can run it. Nothing is written to Active Directory.

.PARAMETER Threshold
    How alike two names must be to count as a near duplicate, from 0 to 1.
    1 means identical. Default 0.85. When not passed, you are asked for it.

.PARAMETER SearchBase
    Distinguished name of an OU to scan, for example "OU=Users,DC=contoso,DC=com".
    When left empty you are asked to choose: whole domain, or pick one OU.

.PARAMETER OutputPath
    CSV report path. Default: DuplicateNames_<timestamp>.csv in the Output Report folder.

.PARAMETER EnabledOnly
    Only scan accounts that are still enabled. Disabled accounts are skipped.

.PARAMETER NoGrid
    Do not open the result window. Useful for scheduled runs.

.EXAMPLE
    .\Find-Duplicate-AD-Names.ps1

.EXAMPLE
    .\Find-Duplicate-AD-Names.ps1 -Threshold 0.9 -EnabledOnly

.EXAMPLE
    .\Find-Duplicate-AD-Names.ps1 -SearchBase "OU=Users,DC=contoso,DC=com" -NoGrid
#>

[CmdletBinding()]
param(
    [ValidateRange(0.0, 1.0)]
    [double]$Threshold = 0.85,

    [string]$SearchBase,

    [string]$OutputPath,

    [switch]$EnabledOnly,

    [switch]$NoGrid
)

$ErrorActionPreference = 'Stop'

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

function Get-NameKey {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $k = $Name -replace '[^\p{L}\p{N} ]', ' '
    $k = ($k -replace '\s+', ' ').Trim().ToLowerInvariant()
    return $k
}

function Get-SortedWordKey {
    param([string]$Key)
    if ($Key -eq '') { return '' }
    $words = $Key.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
    [array]::Sort($words)
    return ($words -join ' ')
}

function Get-EditDistance {
    param([string]$A, [string]$B)
    $n = $A.Length
    $m = $B.Length
    if ($n -eq 0) { return $m }
    if ($m -eq 0) { return $n }
    $prev = New-Object 'int[]' ($m + 1)
    $cur = New-Object 'int[]' ($m + 1)
    for ($j = 0; $j -le $m; $j++) { $prev[$j] = $j }
    for ($i = 1; $i -le $n; $i++) {
        $cur[0] = $i
        for ($j = 1; $j -le $m; $j++) {
            $cost = 1
            if ($A[$i - 1] -eq $B[$j - 1]) { $cost = 0 }
            $del = $prev[$j] + 1
            $ins = $cur[$j - 1] + 1
            $sub = $prev[$j - 1] + $cost
            $cur[$j] = [Math]::Min($del, [Math]::Min($ins, $sub))
        }
        $tmp = $prev; $prev = $cur; $cur = $tmp
    }
    return $prev[$m]
}

function Get-Similarity {
    param([string]$A, [string]$B)
    $max = [Math]::Max($A.Length, $B.Length)
    if ($max -eq 0) { return 1.0 }
    $d = Get-EditDistance -A $A -B $B
    return [Math]::Round(1.0 - ($d / $max), 3)
}

function Get-AccountStatus {
    param($User)
    if ($User.Enabled) { return 'Enabled' }
    return 'Disabled'
}

# Ask the user which part of the domain to scan. Returns a distinguished name
# (empty string means the whole domain), or $null when the user cancels.
function Select-SearchBase {
    Write-ColorOutput "`n=== Step 1: What to scan ===" 'Cyan'
    $choice = Read-Host "Scan the WHOLE domain? (Y = whole domain, N = pick one OU)"

    if ($choice -notmatch '^[Nn]') {
        Write-ColorOutput 'Scanning the whole domain.' 'Green'
        return ''
    }

    Write-ColorOutput 'Retrieving all Organizational Units from Active Directory... (this may take a moment)' 'DarkYellow'
    $allOUs = Get-ADOrganizationalUnit -Filter * -Properties CanonicalName, Description |
        Select-Object Name, CanonicalName, Description, DistinguishedName |
        Sort-Object CanonicalName

    if (-not $allOUs) {
        Write-ColorOutput 'No OUs found. Scanning the whole domain instead.' 'Yellow'
        return ''
    }

    Write-ColorOutput ("Found {0} OUs. Select ONE OU and click OK. Click Cancel to scan the whole domain." -f @($allOUs).Count) 'White'
    $selected = $allOUs | Out-GridView -Title 'Select the OU to scan for duplicate names' -OutputMode Single

    if ($selected) {
        Write-ColorOutput ("Selected OU: {0}" -f $selected.CanonicalName) 'Green'
        return $selected.DistinguishedName
    }

    Write-ColorOutput 'No OU selected - scanning the whole domain.' 'Yellow'
    return ''
}

# ============================================
# MAIN SCRIPT EXECUTION
# ============================================

Write-ColorOutput "`n=================================================================" 'Cyan'
Write-ColorOutput '     AD Duplicate and Similar Name Finder' 'Cyan'
Write-ColorOutput '     Version 1.0 - Read-only scan' 'Cyan'
Write-ColorOutput '=================================================================' 'Cyan'
Write-ColorOutput "Started at: $(Get-Date)" 'Cyan'

# ---------- check the module ----------

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-ColorOutput 'The ActiveDirectory module is missing.' 'Red'
    Write-ColorOutput 'On a server: Install-WindowsFeature RSAT-AD-PowerShell' 'Yellow'
    Write-ColorOutput 'On a Windows PC: install RSAT (Active Directory Domain Services tools).' 'Yellow'
    Write-ColorOutput "`n" 'White'
    do { $exitChoice = Read-Host "Type 'EXIT' to close the script" } while ($exitChoice -ne 'EXIT')
    return
}

Import-Module ActiveDirectory -ErrorAction Stop

# ---------- choose the scope ----------

$scanWholeDomain = $true

if (-not $SearchBase) {
    $picked = Select-SearchBase
    if ($null -eq $picked) {
        Write-ColorOutput 'Scan cancelled by user.' 'Yellow'
        do { $exitChoice = Read-Host "Type 'EXIT' to close the script" } while ($exitChoice -ne 'EXIT')
        return
    }
    $SearchBase = $picked
}

if ($SearchBase) { $scanWholeDomain = $false }

if (-not $scanWholeDomain) {
    if (-not (Get-ADObject -Identity $SearchBase -ErrorAction SilentlyContinue)) {
        Write-ColorOutput ("SearchBase not found in AD: {0}" -f $SearchBase) 'Red'
        do { $exitChoice = Read-Host "Type 'EXIT' to close the script" } while ($exitChoice -ne 'EXIT')
        return
    }
    Write-ColorOutput ("Scope: {0}" -f $SearchBase) 'Cyan'
}
else {
    Write-ColorOutput 'Scope: whole domain' 'Cyan'
}

# ---------- threshold ----------

if (-not $PSBoundParameters.ContainsKey('Threshold')) {
    $answer = Read-Host ("Near-duplicate threshold 0 to 1 (press Enter for {0})" -f $Threshold)
    if (-not [string]::IsNullOrWhiteSpace($answer)) {
        $parsed = 0.0
        if ([double]::TryParse($answer, [ref]$parsed) -and $parsed -ge 0 -and $parsed -le 1) {
            $Threshold = $parsed
        }
        else {
            Write-ColorOutput ("'{0}' is not a number between 0 and 1 - keeping {1}." -f $answer, $Threshold) 'Yellow'
        }
    }
}
Write-ColorOutput ("Threshold: {0}" -f $Threshold) 'Green'

# ---------- output path ----------

if (-not $OutputPath) {
    $folder = $PSScriptRoot
    if (-not $folder) { $folder = (Get-Location).Path }
    $folder = Join-Path $folder 'Output Report'
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    $OutputPath = Join-Path $folder ('DuplicateNames_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}

# ---------- load users (read-only) ----------

Write-ColorOutput "`nLoading users from Active Directory..." 'Cyan'

$adParams = @{
    Filter     = '*'
    Properties = 'displayName', 'sAMAccountName', 'userPrincipalName', 'mail', 'Enabled', 'DistinguishedName', 'whenCreated'
}
if (-not $scanWholeDomain) { $adParams['SearchBase'] = $SearchBase }

$rawUsers = @(Get-ADUser @adParams)
if ($EnabledOnly) {
    $rawUsers = @($rawUsers | Where-Object { $_.Enabled })
}

Write-ColorOutput ("Loaded {0} user accounts." -f $rawUsers.Count) 'Cyan'

$items = foreach ($u in $rawUsers) {
    $key = Get-NameKey -Name $u.displayName
    if ($key -eq '') { continue }
    [pscustomobject]@{
        Name    = $u.displayName
        Key     = $key
        Sorted  = Get-SortedWordKey -Key $key
        Sam     = $u.sAMAccountName
        UPN     = $u.userPrincipalName
        Mail    = $u.mail
        Status  = Get-AccountStatus -User $u
        DN      = $u.DistinguishedName
        Created = $u.whenCreated
    }
}
$items = @($items)

if ($items.Count -eq 0) {
    Write-ColorOutput "`nNo user accounts with a display name were found." 'Green'
    Write-ColorOutput 'Read-only scan finished. No users were changed.' 'Green'
    Write-ColorOutput "`n" 'White'
    do { $exitChoice = Read-Host "Type 'EXIT' to close the script" } while ($exitChoice -ne 'EXIT')
    return
}

$report = New-Object System.Collections.Generic.List[object]
$seenPairs = @{}

function Add-PairRow {
    param($A, $B, [string]$Type, [double]$Similarity)

    $k1 = '{0}|{1}' -f $A.Sam, $B.Sam
    $k2 = '{0}|{1}' -f $B.Sam, $A.Sam
    if ($seenPairs.ContainsKey($k1) -or $seenPairs.ContainsKey($k2)) { return }
    $seenPairs[$k1] = $true

    $report.Add([pscustomobject]@{
            Type       = $Type
            Similarity = $Similarity
            NameA      = $A.Name
            SamA       = $A.Sam
            StatusA    = $A.Status
            NameB      = $B.Name
            SamB       = $B.Sam
            StatusB    = $B.Status
            UPN        = ('{0} / {1}' -f $A.UPN, $B.UPN)
        })
}

# ---------- exact duplicates ----------

Write-ColorOutput "`nChecking for exact duplicate names..." 'Cyan'
$exactGroups = @($items | Group-Object -Property Key | Where-Object { $_.Count -gt 1 })

foreach ($g in $exactGroups) {
    $list = @($g.Group)
    # One row per group, naming every account that shares the name.
    $report.Add([pscustomobject]@{
            Type       = 'Exact'
            Similarity = 1.0
            NameA      = $list[0].Name
            SamA       = (($list | ForEach-Object { $_.Sam }) -join '; ')
            StatusA    = (($list | ForEach-Object { $_.Status }) -join '; ')
            NameB      = ('{0} accounts share this name' -f $list.Count)
            SamB       = ''
            StatusB    = ''
            UPN        = (($list | ForEach-Object { $_.UPN }) -join '; ')
        })
}

# ---------- same words, different order ----------

Write-ColorOutput 'Checking for names with the same words in a different order...' 'Cyan'
$reorderedGroups = @($items | Group-Object -Property Sorted | Where-Object {
        $_.Count -gt 1 -and (@($_.Group | Select-Object -ExpandProperty Key -Unique).Count -gt 1)
    })

foreach ($g in $reorderedGroups) {
    $list = @($g.Group | Group-Object -Property Key | ForEach-Object { $_.Group[0] })
    for ($i = 0; $i -lt $list.Count; $i++) {
        for ($j = $i + 1; $j -lt $list.Count; $j++) {
            Add-PairRow -A $list[$i] -B $list[$j] -Type 'Reordered' -Similarity 1.0
        }
    }
}

# ---------- near duplicates ----------
# Names are only compared inside blocks that share the same first letter, to keep it fast.
# A typo in the very first letter is not caught by this scan.

Write-ColorOutput 'Checking for similar (close) names...' 'Cyan'
$uniqueItems = @($items | Group-Object -Property Key | ForEach-Object { $_.Group[0] })
$blocks = @($uniqueItems | Group-Object -Property { $_.Key.Substring(0, 1) })

foreach ($block in $blocks) {
    $list = @($block.Group)
    for ($i = 0; $i -lt $list.Count; $i++) {
        for ($j = $i + 1; $j -lt $list.Count; $j++) {
            $a = $list[$i]
            $b = $list[$j]
            if ([Math]::Abs($a.Key.Length - $b.Key.Length) -gt 4) { continue }

            $s1 = Get-Similarity -A $a.Key -B $b.Key
            $s2 = Get-Similarity -A $a.Sorted -B $b.Sorted
            $best = [Math]::Max($s1, $s2)

            if ($best -ge $Threshold -and $best -lt 1.0) {
                Add-PairRow -A $a -B $b -Type 'Near' -Similarity $best
            }
        }
    }
}

# ---------- output ----------

$sorted = @($report | Sort-Object -Property @{Expression = 'Type'; Descending = $false }, @{Expression = 'Similarity'; Descending = $true })

Write-ColorOutput "`n=================================================================" 'Cyan'
Write-ColorOutput '     RESULTS' 'Cyan'
Write-ColorOutput '=================================================================' 'Cyan'

if ($sorted.Count -eq 0) {
    Write-ColorOutput 'No duplicate or near-duplicate names found.' 'Green'
}
else {
    $sorted | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    $exactCount = @($sorted | Where-Object { $_.Type -eq 'Exact' }).Count
    $reorderCount = @($sorted | Where-Object { $_.Type -eq 'Reordered' }).Count
    $nearCount = @($sorted | Where-Object { $_.Type -eq 'Near' }).Count

    Write-ColorOutput ("Exact duplicate names      : {0}" -f $exactCount) 'Yellow'
    Write-ColorOutput ("Same words, reordered      : {0}" -f $reorderCount) 'Yellow'
    Write-ColorOutput ("Similar (close) name pairs : {0}" -f $nearCount) 'Yellow'
    Write-ColorOutput ("Report saved to: {0}" -f $OutputPath) 'Cyan'
    Write-ColorOutput "`n" 'White'

    $sorted | Select-Object -First 25 | Format-Table Type, Similarity, NameA, SamA, StatusA, NameB, SamB, StatusB -AutoSize
    if ($sorted.Count -gt 25) {
        Write-ColorOutput '(First 25 shown - see the CSV for all rows.)' 'DarkGray'
    }

    if (-not $NoGrid) {
        $sorted | Select-Object Type, Similarity, NameA, SamA, StatusA, NameB, SamB, StatusB, UPN |
            Out-GridView -Title 'Duplicate and similar display names (read-only, nothing was changed)'
    }
}

Write-ColorOutput "`n=================================================================" 'Cyan'
Write-ColorOutput "     Script Completed at: $(Get-Date)" 'Cyan'
Write-ColorOutput '=================================================================' 'Cyan'
Write-ColorOutput 'Read-only scan finished. No users were changed.' 'Green'

# ============================================
# MANUAL EXIT TRIGGER
# ============================================

Write-ColorOutput "`n" 'White'
do {
    $exitChoice = Read-Host "Type 'EXIT' to close the script"
} while ($exitChoice -ne 'EXIT')
