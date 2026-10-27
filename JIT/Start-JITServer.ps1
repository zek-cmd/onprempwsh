#Requires -Version 7.0
<#
.SYNOPSIS
    Serves the JIT web page on http://localhost:40000/

.DESCRIPTION
    Small local web server. It serves index.html and answers:
      GET  /api/health       server status
      GET  /api/ous          list of OUs in the domain
      POST /api/computers    computers inside a chosen OU
      POST /api/ping         which computers are online
      POST /api/jit          Grant / Revoke / Status (built into this script)
      GET  /api/log          session log (table)
    Every grant, revoke and status action is written to logs\session-log.json.
    Listens on localhost only, so other PCs cannot reach it.

    Run as Administrator (needed for WinRM, group changes and AD queries).
    Needs the ActiveDirectory module (RSAT) for OU and computer lists.

.EXAMPLE
    .\Start-JITServer.ps1
    Then open http://localhost:40000/ in your browser.
#>
[CmdletBinding()]
param(
    [int]$Port = 40000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$webRoot   = $PSScriptRoot
$logDir    = Join-Path $PSScriptRoot 'logs'
$logFile   = Join-Path $logDir 'session-log.json'
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$hasAD = [bool](Get-Module -ListAvailable -Name ActiveDirectory)
if ($hasAD) { Import-Module ActiveDirectory -ErrorAction SilentlyContinue }

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()
Write-Host "JIT web page running at http://localhost:$Port/  (Ctrl+C to stop)" -ForegroundColor Cyan

function Send-Response {
    param($Response, [int]$Code, [string]$Body, [string]$Type = 'application/json; charset=utf-8')
    $bytes = [Text.Encoding]::UTF8.GetBytes($Body)
    $Response.StatusCode  = $Code
    $Response.ContentType = $Type
    $Response.ContentLength64 = $bytes.Length
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Response.OutputStream.Close()
}

function Read-Payload {
    param($Request)
    $reader = [IO.StreamReader]::new($Request.InputStream, $Request.ContentEncoding)
    return $reader.ReadToEnd() | ConvertFrom-Json
}

function Read-SessionLog {
    if (-not (Test-Path $logFile)) { return @() }
    $raw = Get-Content -LiteralPath $logFile -Raw
    if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
    return @($raw | ConvertFrom-Json)
}

function Add-SessionLog {
    param([string]$Action, [string]$User, [string]$Computer, [string]$Status, [string]$Detail)
    $entries = [System.Collections.Generic.List[object]]::new()
    foreach ($e in (Read-SessionLog)) { $entries.Add($e) }
    $entries.Add([pscustomobject]@{
        Time     = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Operator = $env:USERNAME
        Action   = $Action
        User     = $User
        Computer = $Computer
        Status   = $Status
        Detail   = $Detail
    })
    # Keep the newest 1000 entries.
    $keep = @($entries | Select-Object -Last 1000)
    ConvertTo-Json -InputObject $keep -Depth 4 | Set-Content -LiteralPath $logFile -Encoding UTF8
}

function Get-OUList {
    if (-not $hasAD) { throw 'ActiveDirectory module is not installed (install RSAT).' }
    Get-ADOrganizationalUnit -Filter * -Properties Name |
        Sort-Object DistinguishedName |
        ForEach-Object { [pscustomobject]@{ Name = $_.Name; DN = $_.DistinguishedName } }
}

function Get-OUComputers {
    param([string]$OUDN)
    if (-not $hasAD) { throw 'ActiveDirectory module is not installed (install RSAT).' }
    Get-ADComputer -SearchBase $OUDN -SearchScope Subtree -Filter * |
        Sort-Object Name |
        ForEach-Object { $_.Name }
}

function Get-LoggedOnUsers {
    param([string[]]$Names)
    # Ask each PC who is logged on, in parallel, over WinRM (PowerShell 7).
    $Names | ForEach-Object -Parallel {
        $name = $_
        try {
            $user = Invoke-Command -ComputerName $name -ScriptBlock {
                (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
            } -ErrorAction Stop
            [pscustomobject]@{ Computer = $name; User = [string]$user; Status = $(if ($user) { 'OK' } else { 'NoUser' }) }
        }
        catch {
            [pscustomobject]@{ Computer = $name; User = ''; Status = 'Failed' }
        }
    } -ThrottleLimit 32
}

# ----------------------------------------------- JIT logic (runs on target PCs) ---

$remoteGrant = {
    param([string]$Account, [int]$Minutes, [string]$TaskName)

    $group = $null
    try { $group = (Get-LocalGroup -SID 'S-1-5-32-544' -ErrorAction Stop).Name } catch { }
    if (-not $group) { $group = 'Administrators' }

    # 1. Add the account to local Administrators (skip if already a member).
    $already = Get-LocalGroupMember -Group $group -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ieq $Account }
    if (-not $already) {
        Add-LocalGroupMember -Group $group -Member $Account
    }

    $loggedOn = $null
    try { $loggedOn = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName } catch { }

    # 2. Timer: a one-time scheduled task that removes the account later.
    $removeCmd = "Remove-LocalGroupMember -Group '$group' -Member '$Account' -ErrorAction SilentlyContinue; " +
                 "Unregister-ScheduledTask -TaskName '$TaskName' -Confirm:`$false -ErrorAction SilentlyContinue"

    # /TR is limited to 261 characters, so the code goes into a small script file
    # on the PC, and the task only runs that file.
    $jitDir = Join-Path $env:ProgramData 'JIT'
    if (-not (Test-Path $jitDir)) { New-Item -ItemType Directory -Path $jitDir -Force | Out-Null }
    $removeFile = Join-Path $jitDir ($TaskName + '-remove.ps1')
    Set-Content -LiteralPath $removeFile -Value $removeCmd -Encoding UTF8
    $runArg  = "powershell.exe -NoProfile -WindowStyle Hidden -File `"$removeFile`""
    $when    = (Get-Date).AddMinutes($Minutes)

    # Replace any older timer for the same account, then create a one-time task
    # with schtasks.exe (fixed date and time, no trigger object, no XML).
    schtasks.exe /Delete /TN $TaskName /F 2>$null | Out-Null
    # Only the time is passed (no /SD date), so the target PC's regional date
    # format cannot cause "Incorrect Start Date". The timer is for today.
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $st  = $when.ToString('HH:mm', $inv)
    schtasks.exe /Create /TN $TaskName /TR $runArg /SC ONCE /ST $st /RU SYSTEM /RL HIGHEST /F | Out-Null

    # 3. Mark the session: a task that runs at shutdown and removes the admin
    #    rights and the timer. Shutting the PC down therefore ends the session.
    $cleanCmd = "Remove-LocalGroupMember -Group '$group' -Member '$Account' -ErrorAction SilentlyContinue; " +
                "Unregister-ScheduledTask -TaskName '$TaskName' -Confirm:`$false -ErrorAction SilentlyContinue; " +
                "Unregister-ScheduledTask -TaskName '$TaskName-Shutdown' -Confirm:`$false -ErrorAction SilentlyContinue"
    $cleanFile = Join-Path $jitDir ($TaskName + '-clean.ps1')
    Set-Content -LiteralPath $cleanFile -Value $cleanCmd -Encoding UTF8
    # Startup trigger (no event trigger, no hand-written XML): after any shutdown
    # or restart the PC boots, and this task clears the session right away.
    $cleanRun = "powershell.exe -NoProfile -WindowStyle Hidden -File `"$cleanFile`""
    schtasks.exe /Delete /TN ($TaskName + '-Shutdown') /F 2>$null | Out-Null
    schtasks.exe /Create /TN ($TaskName + '-Shutdown') /TR $cleanRun /SC ONSTART /RU SYSTEM /RL HIGHEST /F | Out-Null

    [pscustomobject]@{
        Computer     = $env:COMPUTERNAME
        Account      = $Account
        Granted      = $true
        ExpiresAt    = (Get-Date).AddMinutes($Minutes).ToString('yyyy-MM-dd HH:mm:ss')
        LoggedOnUser = $loggedOn
        SessionTask  = $true
    }
}

$remoteRevoke = {
    param([string]$Account, [string]$TaskName)

    $group = $null
    try { $group = (Get-LocalGroup -SID 'S-1-5-32-544' -ErrorAction Stop).Name } catch { }
    if (-not $group) { $group = 'Administrators' }
    Remove-LocalGroupMember -Group $group -Member $Account -ErrorAction SilentlyContinue
    schtasks.exe /Delete /TN $TaskName /F 2>$null | Out-Null
    schtasks.exe /Delete /TN ($TaskName + '-Shutdown') /F 2>$null | Out-Null

    $loggedOn = $null
    try { $loggedOn = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName } catch { }

    [pscustomobject]@{
        Computer     = $env:COMPUTERNAME
        Account      = $Account
        Granted      = $false
        ExpiresAt    = $null
        LoggedOnUser = $loggedOn
        SessionTask  = $false
    }
}

$remoteStatus = {
    param([string]$Account, [string]$TaskName)

    $group = $null
    try { $group = (Get-LocalGroup -SID 'S-1-5-32-544' -ErrorAction Stop).Name } catch { }
    if (-not $group) { $group = 'Administrators' }
    $member = Get-LocalGroupMember -Group $group -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ieq $Account }
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    $expires = $null
    if ($task) {
        $info = Get-ScheduledTaskInfo -TaskName $TaskName
        $expires = $info.NextRunTime
    }

    $loggedOn = $null
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem
        $loggedOn = $cs.UserName
    } catch { }

    [pscustomobject]@{
        Computer     = $env:COMPUTERNAME
        Account      = $Account
        Granted      = [bool]$member
        ExpiresAt    = $expires
        LoggedOnUser = $loggedOn
        SessionTask  = [bool]$task
    }
}

function Invoke-JIT {
    param([string]$User, [string[]]$Computer, [string]$Action, [int]$Minutes = 60, [string]$Domain)

    if (-not $Domain) {
        # Dynamic domain detection: read the domain this server PC is joined to.
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem
        if (-not $cs.PartOfDomain) { throw 'This computer is not joined to a domain. Enter a domain name.' }
        $Domain = $cs.Domain
    }
    $netBios = ($Domain -split '\.')[0].ToUpperInvariant()
    $account = if ($User -match '\\') { $User } else { "$netBios\$User" }
    $taskName = 'JIT-Revoke-' + ($account -replace '[\\\/ ]', '_')

    $remoteBlock = switch ($Action) {
        'Grant'  { $remoteGrant }
        'Revoke' { $remoteRevoke }
        default  { $remoteStatus }
    }
    $remoteArgs = if ($Action -eq 'Grant') { @($account, $Minutes, $taskName) } else { @($account, $taskName) }

    # ForEach-Object -Parallel cannot receive a script block variable, so the
    # action is passed as text and rebuilt inside each parallel run.
    $remoteText = $remoteBlock.ToString()

    $Computer | ForEach-Object -Parallel {
        $name = $_
        $block = [scriptblock]::Create($using:remoteText)
        $remArgs = $using:remoteArgs

        $online = Test-Connection -TargetName $name -Count 1 -Quiet -TimeoutSeconds 2
        if (-not $online) {
            return [pscustomobject]@{ Computer = $name; Status = 'Offline'; Detail = 'No ping reply' }
        }
        try {
            $r = Invoke-Command -ComputerName $name -ScriptBlock $block -ArgumentList $remArgs -ErrorAction Stop
            [pscustomobject]@{ Computer = $name; Status = 'OK'; Detail = $r }
        }
        catch {
            [pscustomobject]@{ Computer = $name; Status = 'Failed'; Detail = $_.Exception.Message }
        }
    } -ThrottleLimit 10
}

function Test-ComputersOnline {
    param([string[]]$Names)
    # Ping all at once in parallel (PowerShell 7).
    $Names | ForEach-Object -Parallel {
        $ok = Test-Connection -TargetName $_ -Count 1 -Quiet -TimeoutSeconds 2
        [pscustomobject]@{ Computer = $_; Online = [bool]$ok }
    } -ThrottleLimit 32
}

try {
    while ($listener.IsListening) {
        $ctx  = $listener.GetContext()
        $req  = $ctx.Request
        $res  = $ctx.Response
        $path = $req.Url.AbsolutePath

        try {
            if ($req.HttpMethod -eq 'GET' -and $path -eq '/') {
                $html = Get-Content -LiteralPath (Join-Path $webRoot 'index.html') -Raw
                Send-Response $res 200 $html 'text/html; charset=utf-8'
            }
            elseif ($req.HttpMethod -eq 'GET' -and $path -eq '/api/health') {
                $body = [pscustomobject]@{ ok = $true; host = $env:COMPUTERNAME; ad = $hasAD } | ConvertTo-Json -Compress
                Send-Response $res 200 $body
            }
            elseif ($req.HttpMethod -eq 'GET' -and $path -eq '/api/ous') {
                $ous = @(Get-OUList)
                Send-Response $res 200 ([pscustomobject]@{ ous = $ous } | ConvertTo-Json -Depth 4 -Compress)
            }
            elseif ($req.HttpMethod -eq 'POST' -and $path -eq '/api/computers') {
                $payload = Read-Payload $req
                $names = @(Get-OUComputers -OUDN ([string]$payload.ou))
                Send-Response $res 200 ([pscustomobject]@{ computers = $names } | ConvertTo-Json -Depth 4 -Compress)
            }
            elseif ($req.HttpMethod -eq 'POST' -and $path -eq '/api/detect') {
                $payload = Read-Payload $req
                $list = @(Get-LoggedOnUsers -Names @($payload.computers))
                Send-Response $res 200 ([pscustomobject]@{ results = $list } | ConvertTo-Json -Depth 4 -Compress)
            }
            elseif ($req.HttpMethod -eq 'POST' -and $path -eq '/api/ping') {
                $payload = Read-Payload $req
                $list = @(Test-ComputersOnline -Names @($payload.computers))
                Send-Response $res 200 ([pscustomobject]@{ results = $list } | ConvertTo-Json -Depth 4 -Compress)
            }
            elseif ($req.HttpMethod -eq 'POST' -and $path -eq '/api/jit') {
                $payload = Read-Payload $req

                $params = @{
                    User     = [string]$payload.user
                    Computer = [string[]]@($payload.computers)
                    Action   = switch ([string]$payload.action) {
                        'Grant'  { 'Grant' }
                        'Revoke' { 'Revoke' }
                        default  { 'Status' }
                    }
                }
                if ($payload.minutes) { $params.Minutes = [int]$payload.minutes }
                if ($payload.domain)  { $params.Domain  = [string]$payload.domain }

                $results = @(Invoke-JIT @params)

                foreach ($r in $results) {
                    $detail = if ($r.Status -eq 'OK') {
                        if ($r.Detail.Granted) { "Granted until $($r.Detail.ExpiresAt)" } else { 'Standard (no admin)' }
                    } else { [string]$r.Detail }
                    if ($r.Status -eq 'OK' -and $r.Detail.LoggedOnUser) {
                        $detail += " | Logged on: $($r.Detail.LoggedOnUser)"
                    }
                    Add-SessionLog -Action $params.Action -User $params.User -Computer $r.Computer `
                                   -Status $r.Status -Detail $detail
                }

                $body = [pscustomobject]@{ results = $results } | ConvertTo-Json -Depth 6 -Compress
                Send-Response $res 200 $body
            }
            elseif ($req.HttpMethod -eq 'GET' -and $path -eq '/api/log') {
                $entries = @(Read-SessionLog | Select-Object -Last 200)
                Send-Response $res 200 ([pscustomobject]@{ entries = $entries } | ConvertTo-Json -Depth 4 -Compress)
            }
            elseif ($req.HttpMethod -eq 'POST' -and $path -eq '/api/log/clear') {
                # Empty the session log file (the file is kept, just with no records).
                ConvertTo-Json -InputObject @() | Set-Content -LiteralPath $logFile -Encoding UTF8
                Send-Response $res 200 '{"ok":true}'
            }
            else {
                Send-Response $res 404 '{"error":"Not found"}'
            }
        }
        catch {
            $err = [pscustomobject]@{ error = $_.Exception.Message } | ConvertTo-Json -Compress
            Send-Response $res 500 $err
        }
    }
}
finally {
    $listener.Stop()
    $listener.Close()
}
