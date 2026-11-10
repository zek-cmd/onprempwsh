# ADUserGuard.ps1
# Shared match checking and per-user confirmation for the ONPREM user-account scripts.
# It does two things:
#   1) Get-ADUserMatch / Show-ADUserMatch - look the person up in Active Directory
#      and flag an exact match or a very similar name (typos, duplicates).
#   2) Confirm-UserAction - ask a plain yes/no question about that one user
#      before anything is created, moved, reset or deleted.
# Dot-source it from a script in the same folder:
#     . (Join-Path $PSScriptRoot 'ADUserGuard.ps1')
# PowerShell 5.1 compatible.

# ---------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------

# Turn a distinguished name into the OU path (drop the CN= part)
function Get-ADMatchLocation {
    param([string]$DistinguishedName)

    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return "" }
    return ($DistinguishedName -replace '^CN=.*?,(?=OU=|DC=|CN=)', '')
}

# Similarity between two names, 0.0 (nothing alike) to 1.0 (identical)
function Get-NameSimilarity {
    param(
        [string]$Name1,
        [string]$Name2
    )

    if ([string]::IsNullOrWhiteSpace($Name1) -or [string]::IsNullOrWhiteSpace($Name2)) { return 0.0 }

    $a = ($Name1.Trim() -replace '\s+', ' ').ToLower()
    $b = ($Name2.Trim() -replace '\s+', ' ').ToLower()

    if ($a -ceq $b) { return 1.0 }

    $len1 = $a.Length
    $len2 = $b.Length
    if ($len1 -eq 0 -or $len2 -eq 0) { return 0.0 }

    $previous = New-Object 'int[]' ($len2 + 1)
    $current = New-Object 'int[]' ($len2 + 1)

    for ($j = 0; $j -le $len2; $j++) { $previous[$j] = $j }

    for ($i = 1; $i -le $len1; $i++) {
        $current[0] = $i
        $char1 = $a[$i - 1]

        for ($j = 1; $j -le $len2; $j++) {
            $char2 = $b[$j - 1]
            $cost = 1
            if ($char1 -ceq $char2) { $cost = 0 }

            $deletion = $previous[$j] + 1
            $insertion = $current[$j - 1] + 1
            $substitution = $previous[$j - 1] + $cost

            $min = $deletion
            if ($insertion -lt $min) { $min = $insertion }
            if ($substitution -lt $min) { $min = $substitution }

            $current[$j] = $min
        }

        $temp = $previous
        $previous = $current
        $current = $temp
    }

    $maxLength = [Math]::Max($len1, $len2)
    if ($maxLength -eq 0) { return 0.0 }

    return [Math]::Round((1.0 - ($previous[$len2] / $maxLength)), 4)
}

# ---------------------------------------------------------------
# Match checking
# ---------------------------------------------------------------

# Look for accounts that are the same person or a close name.
# Returns an object with:
#   Exact      - accounts matching the name or username exactly
#   Close      - accounts with a similar name (likely the same person / a duplicate)
#   MatchCount - how many were found
#   HadError   - true when AD could not be searched
function Get-ADUserMatch {
    param(
        [string]$FullName,
        [string]$Username,
        [string]$SearchBase,
        [string]$ExcludeSamAccountName,
        [double]$Threshold = 0.80,
        [int]$MaxResults = 5
    )

    $state = [PSCustomObject]@{
        Exact      = @()
        Close      = @()
        MatchCount = 0
        HadError   = $false
    }

    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        $state.HadError = $true
        return $state
    }

    $props = @('Name', 'DisplayName', 'SamAccountName', 'DistinguishedName', 'Enabled')
    $exact = @()
    $seen = @{}

    # --- exact username ---
    if (-not [string]::IsNullOrWhiteSpace($Username)) {
        $escapedUser = $Username.Trim() -replace "'", "''"
        try {
            $byUser = @(Get-ADUser -Filter "SamAccountName -eq '$escapedUser'" -Properties $props -ErrorAction Stop)
            foreach ($u in $byUser) {
                if (-not $seen.ContainsKey($u.DistinguishedName)) {
                    $seen[$u.DistinguishedName] = $true
                    $exact += $u
                }
            }
        }
        catch {
            $state.HadError = $true
        }
    }

    # --- exact display name and exact CN name ---
    if (-not [string]::IsNullOrWhiteSpace($FullName)) {
        $escapedName = $FullName.Trim() -replace "'", "''"

        foreach ($filter in @("Name -eq '$escapedName'", "DisplayName -eq '$escapedName'")) {
            try {
                $byName = @(Get-ADUser -Filter $filter -Properties $props -ErrorAction Stop)
                foreach ($u in $byName) {
                    if (-not $seen.ContainsKey($u.DistinguishedName)) {
                        $seen[$u.DistinguishedName] = $true
                        $exact += $u
                    }
                }
            }
            catch {
                $state.HadError = $true
            }
        }
    }

    # --- close names: search each meaningful word of the name, then score it ---
    if (-not [string]::IsNullOrWhiteSpace($FullName)) {
        $tokens = @($FullName.Trim() -split '\s+' | Where-Object { $_ -and $_.Length -ge 3 })
        $candidates = @()

        foreach ($token in $tokens) {
            $escapedToken = $token -replace "'", "''"
            try {
                $found = @(Get-ADUser -Filter "Name -like '*$escapedToken*'" -Properties $props -ResultSetSize 500 -ErrorAction Stop)
                $candidates += $found
            }
            catch {
                $state.HadError = $true
            }
        }

        $scored = @()
        if ($candidates.Count -gt 0) {
            $candidateList = @($candidates | Sort-Object DistinguishedName -Unique)

            foreach ($candidate in $candidateList) {
                if ($seen.ContainsKey($candidate.DistinguishedName)) { continue }

                $scoreName = Get-NameSimilarity -Name1 $FullName -Name2 $candidate.Name
                $scoreDisplay = Get-NameSimilarity -Name1 $FullName -Name2 $candidate.DisplayName
                $best = [Math]::Max($scoreName, $scoreDisplay)

                if ($best -ge $Threshold) {
                    $scored += [PSCustomObject]@{
                        Name              = $candidate.Name
                        DisplayName       = $candidate.DisplayName
                        SamAccountName    = $candidate.SamAccountName
                        DistinguishedName = $candidate.DistinguishedName
                        Enabled           = $candidate.Enabled
                        CurrentOU         = (Get-ADMatchLocation $candidate.DistinguishedName)
                        SimilarityPercent = [Math]::Round($best * 100, 1)
                    }
                }
            }
        }

        if ($scored.Count -gt 0) {
            $state.Close = @($scored | Sort-Object SimilarityPercent -Descending | Select-Object -First $MaxResults)
        }
    }

    # Never flag the very account we are working on (moving, resetting, deleting it)
    if (-not [string]::IsNullOrWhiteSpace($ExcludeSamAccountName)) {
        $exact = @($exact | Where-Object { $_.SamAccountName -ne $ExcludeSamAccountName })
        if ($state.Close.Count -gt 0) {
            $state.Close = @($state.Close | Where-Object { $_.SamAccountName -ne $ExcludeSamAccountName })
        }
    }

    $state.Exact = @($exact)
    $state.MatchCount = $state.Exact.Count + $state.Close.Count

    return $state
}

# Print the matches found. Returns $true when something worth a yes/no question was found.
function Show-ADUserMatch {
    param(
        [string]$FullName,
        [string]$Username,
        $Match
    )

    if (-not $Match -or $Match.MatchCount -eq 0) { return $false }

    Write-Host ""
    Write-Host "  MATCH CHECK: found $($Match.MatchCount) account(s) that may be this person" -ForegroundColor Yellow

    foreach ($u in $Match.Exact) {
        $enabledText = if ($u.Enabled) { "enabled" } else { "disabled" }
        Write-Host ("    [EXACT] {0} ({1}) - {2}, {3}" -f $u.Name, $u.SamAccountName, (Get-ADMatchLocation $u.DistinguishedName), $enabledText) -ForegroundColor Red
    }

    foreach ($u in $Match.Close) {
        Write-Host ("    [CLOSE] {0} ({1}) - {2} similar, {3}" -f $u.Name, $u.SamAccountName, "$($u.SimilarityPercent)%", (Get-ADMatchLocation $u.DistinguishedName)) -ForegroundColor Yellow
    }

    return $true
}

# ---------------------------------------------------------------
# Confirmation
# ---------------------------------------------------------------

# Ask a plain yes/no question about one user. Returns $true for yes, $false for no.
function Confirm-UserAction {
    param(
        [Parameter(Mandatory = $true)][string]$Action,
        [Parameter(Mandatory = $true)][string]$TargetName,
        [string]$Username,
        [string]$Detail,
        [string]$Default = 'N'
    )

    Write-Host ""
    Write-Host "  ------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ("  About to {0}: {1}" -f $Action, $TargetName) -ForegroundColor White
    if (-not [string]::IsNullOrWhiteSpace($Username)) {
        Write-Host ("    Username : {0}" -f $Username) -ForegroundColor Gray
    }
    if (-not [string]::IsNullOrWhiteSpace($Detail)) {
        Write-Host ("    Details  : {0}" -f $Detail) -ForegroundColor Gray
    }

    while ($true) {
        $answer = Read-Host ("  {0} for this user? (y=yes / n=no)" -f $Action)
        if ([string]::IsNullOrWhiteSpace($answer)) { $answer = $Default }

        switch ($answer.Trim().ToLower()) {
            'y' { return $true }
            'yes' { return $true }
            'n' { return $false }
            'no' { return $false }
        }

        Write-Host "  Please type y for yes or n for no - nothing has been done yet." -ForegroundColor Yellow
    }
}
