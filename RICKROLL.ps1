#Requires -Version 7.0
# Remote Rickroll Sender
# Run from Domain Controller with Administrator privileges
# Requires: ActiveDirectory module, PowerShell 7 (parallel sends)
#
# What it does, per target PC:
#   1. Writes two tiny helper scripts into C:\ProgramData\MAS on that PC.
#   2. Registers a one-time task that opens Chrome on the video, in the
#      session of whoever is logged on there (so the user really sees it).
#   3. Registers a second task, running as SYSTEM a couple of minutes later,
#      that deletes the first task and both helper files - the task removes
#      itself after it fired.

Import-Module ActiveDirectory -ErrorAction Stop

# === CONFIGURATION ===
$VideoUrl    = 'https://youtu.be/dQw4w9WgXcQ?si=6TEYyQWpy-qEcbrV'
$TaskName    = 'MAS_Rickroll'
$HelperName  = 'rickroll'
$DelayMinutes = 1     # play the video this many minutes from now
$ThrottleLimit = 20   # how many PCs are worked on at the same time

$domainName = (Get-ADDomain).Name

Clear-Host
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    REMOTE RICKROLL SENDER" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Domain: $domainName" -ForegroundColor Gray
Write-Host "  Video : $VideoUrl" -ForegroundColor Gray
Write-Host "  Plays : $DelayMinutes minute(s) after sending" -ForegroundColor Gray
Write-Host ""

# === SELECT COMPUTERS ===
Write-Host "=== SELECT COMPUTERS ===" -ForegroundColor Cyan
Write-Host "Loading OUs..." -ForegroundColor Yellow

$domainDN = (Get-ADDomain).DistinguishedName
$allOUs = @()

# Domain Root as the first option
$domainComps = @(Get-ADComputer -Filter {Enabled -eq $true} -SearchBase $domainDN -SearchScope OneLevel -ErrorAction SilentlyContinue)
if ($domainComps.Count -gt 0) {
    $allOUs += [PSCustomObject]@{
        Name = "[Domain Root]"
        DistinguishedName = $domainDN
        ComputerCount = $domainComps.Count
    }
}

# Every OU that holds at least one enabled computer
$ous = Get-ADOrganizationalUnit -Filter * | Sort-Object Name
foreach ($ou in $ous) {
    $comps = @(Get-ADComputer -Filter {Enabled -eq $true} -SearchBase $ou.DistinguishedName -SearchScope OneLevel -ErrorAction SilentlyContinue)
    if ($comps.Count -gt 0) {
        $allOUs += [PSCustomObject]@{
            Name = $ou.Name
            DistinguishedName = $ou.DistinguishedName
            ComputerCount = $comps.Count
        }
    }
}

if ($allOUs.Count -eq 0) {
    Write-Host "No OUs with computers found." -ForegroundColor Red
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "    Script Execution Complete" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    if ($Host.Name -eq 'ConsoleHost') {
        Write-Host "Press any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    Exit 0
}

Write-Host "Please select an OU from the popup window..." -ForegroundColor Yellow
$selectedOU = $allOUs | Select-Object Name, @{Name="Computers";Expression={$_.ComputerCount}}, DistinguishedName |
    Out-GridView -Title "Select Target OU (Use Filter to Search)" -OutputMode Single

if (-not $selectedOU) {
    Write-Host "No OU selected." -ForegroundColor Red
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "    Script Execution Complete" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    if ($Host.Name -eq 'ConsoleHost') {
        Write-Host "Press any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    Exit 0
}

Write-Host "`nSelected OU: $($selectedOU.Name)" -ForegroundColor Green
Write-Host "Scanning for online computers..." -ForegroundColor Cyan

$computers = Get-ADComputer -Filter {Enabled -eq $true} -SearchBase $selectedOU.DistinguishedName -Properties DNSHostName, OperatingSystem -ErrorAction SilentlyContinue |
    Sort-Object Name

if ($computers.Count -eq 0) {
    Write-Host "No enabled computers found in this OU." -ForegroundColor Red
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "    Script Execution Complete" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    if ($Host.Name -eq 'ConsoleHost') {
        Write-Host "Press any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    Exit 0
}

Write-Host "Testing connectivity for $($computers.Count) computers..." -ForegroundColor Cyan
Write-Host "This may take a moment..." -ForegroundColor Gray

# All pings run at once (PowerShell 7 parallel)
$pingResults = $computers | ForEach-Object -Parallel {
    $computerName = if ($_.DNSHostName) { $_.DNSHostName } else { $_.Name }
    $isOnline = Test-Connection -TargetName $computerName -Count 1 -Quiet -TimeoutSeconds 2 -ErrorAction SilentlyContinue
    [PSCustomObject]@{
        Name = $_.Name
        FQDN = $computerName
        OperatingSystem = $_.OperatingSystem
        Online = [bool]$isOnline
    }
} -ThrottleLimit 50

$onlineComputers = @()
$offlineCount = 0
foreach ($result in ($pingResults | Sort-Object Name)) {
    Write-Host "  Testing $($result.Name)... " -NoNewline
    if ($result.Online) {
        Write-Host "ONLINE" -ForegroundColor Green
        $onlineComputers += [PSCustomObject]@{
            Name = $result.Name
            FQDN = $result.FQDN
            OperatingSystem = $result.OperatingSystem
        }
    } else {
        Write-Host "OFFLINE" -ForegroundColor Red
        $offlineCount++
    }
}

Write-Host "`nResults: $($onlineComputers.Count) online, $offlineCount offline" -ForegroundColor Cyan

if ($onlineComputers.Count -eq 0) {
    Write-Host "No online computers found in this OU." -ForegroundColor Red
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "    Script Execution Complete" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    if ($Host.Name -eq 'ConsoleHost') {
        Write-Host "Press any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    Exit 0
}

Write-Host "`nPlease select computers from the popup window..." -ForegroundColor Yellow
Write-Host "Use Ctrl+Click for multiple selection" -ForegroundColor Gray

$selectedComputers = $onlineComputers |
    Select-Object Name, OperatingSystem |
    Out-GridView -Title "Select Target Computers (Ctrl+Click for multiple)" -OutputMode Multiple

if (-not $selectedComputers -or $selectedComputers.Count -eq 0) {
    Write-Host "No computers selected." -ForegroundColor Red
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "    Script Execution Complete" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    if ($Host.Name -eq 'ConsoleHost') {
        Write-Host "Press any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    Exit 0
}

$finalSelection = @()
foreach ($selected in $selectedComputers) {
    $comp = $onlineComputers | Where-Object { $_.Name -eq $selected.Name }
    if ($comp) {
        $finalSelection += $comp
    }
}

Write-Host "`n$($finalSelection.Count) computer(s) selected:" -ForegroundColor Green
foreach ($c in $finalSelection) {
    Write-Host "  - $($c.Name) ($($c.OperatingSystem))" -ForegroundColor Gray
}

# === CONFIRM ===
Write-Host "`n=== SEND RICKROLL ===" -ForegroundColor Cyan
Write-Host "Chrome will open on the selected PCs, in the session of the user logged on there." -ForegroundColor Gray
$confirm = Read-Host "Send the rickroll to these $($finalSelection.Count) computer(s)? (Y/N)"

# === WORK THAT RUNS ON EACH TARGET PC ===
$remoteRickroll = {
    param([string]$Url, [string]$TaskName, [string]$HelperName, [int]$DelayMinutes)

    $me = $env:COMPUTERNAME

    # Chrome must open in the session of a real logged-on user, otherwise it
    # runs invisibly in session 0 and nobody sees anything.
    $loggedOn = $null
    try { $loggedOn = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName } catch { }
    if (-not $loggedOn) {
        [pscustomobject]@{
            Computer = $me
            Status   = 'Skipped'
            User     = ''
            FiresAt  = ''
            Detail   = 'No user logged on - the video would be invisible'
        }
        return
    }

    $dir = Join-Path $env:ProgramData 'MAS'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $playFile  = Join-Path $dir ($HelperName + '-play.ps1')
    $cleanFile = Join-Path $dir ($HelperName + '-clean.ps1')
    $cleanTask = $TaskName + '-Cleanup'

    # --- helper 1: opens the video in Chrome (or the default browser) ---
    $playBody = @'
$url = '__URL__'
$candidates = @()
if ($env:ProgramFiles)        { $candidates += (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe') }
if (${env:ProgramFiles(x86)}) { $candidates += (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe') }
if ($env:LOCALAPPDATA)        { $candidates += (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe') }
$chrome = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
try {
    if ($chrome) { Start-Process -FilePath $chrome -ArgumentList '--new-window', $url }
    else         { Start-Process $url }
} catch { }
'@
    $playText = $playBody.Replace('__URL__', $Url.Replace("'", "''"))
    Set-Content -LiteralPath $playFile -Value $playText -Encoding UTF8

    # --- helper 2: removes the play task and both files, then itself ---
    $cleanBody = @'
schtasks.exe /Delete /TN '__PLAYTASK__' /F 2>$null | Out-Null
Remove-Item -LiteralPath '__PLAYFILE__' -Force -ErrorAction SilentlyContinue
schtasks.exe /Delete /TN '__CLEANTASK__' /F 2>$null | Out-Null
Start-Sleep -Seconds 2
Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
'@
    $cleanText = $cleanBody.Replace('__PLAYTASK__', $TaskName).Replace('__PLAYFILE__', $playFile).Replace('__CLEANTASK__', $cleanTask)
    Set-Content -LiteralPath $cleanFile -Value $cleanText -Encoding UTF8

    # --- task XML: a one-time trigger, and whoever should run it ---
    function New-TaskXml {
        param([string]$User, [string]$LogonType, [string]$RunLevel, [string]$StartBoundary, [string]$ScriptPath, [string]$Description)

        $escUser = [System.Security.SecurityElement]::Escape($User)
        $escDesc = [System.Security.SecurityElement]::Escape($Description)
        $arg     = '-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $ScriptPath + '"'
        $escArg  = [System.Security.SecurityElement]::Escape($arg)

        # The task XML schema has no 'ServiceAccount' logon type (the cmdlet does,
        # the XML does not). A service account like SYSTEM just leaves the element out.
        $logonXml = if ($LogonType -and $LogonType -notin @('ServiceAccount', 'None')) {
            "      <LogonType>$LogonType</LogonType>"
        } else { '' }

        return @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>$escDesc</Description>
  </RegistrationInfo>
  <Triggers>
    <TimeTrigger>
      <StartBoundary>$StartBoundary</StartBoundary>
      <Enabled>true</Enabled>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$escUser</UserId>
$logonXml
      <RunLevel>$RunLevel</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>$escArg</Arguments>
    </Exec>
  </Actions>
</Task>
"@
    }

    $inv      = [Globalization.CultureInfo]::InvariantCulture
    $startAt  = (Get-Date).AddMinutes([double]$DelayMinutes)
    $cleanAt  = $startAt.AddMinutes(2)
    $startStr = $startAt.ToString('yyyy-MM-ddTHH:mm:ss', $inv)
    $cleanStr = $cleanAt.ToString('yyyy-MM-ddTHH:mm:ss', $inv)

    try {
        # 1. The play task, in the logged-on user's own session
        $playXml = New-TaskXml -User $loggedOn -LogonType 'InteractiveToken' -RunLevel 'LeastPrivilege' `
            -StartBoundary $startStr -ScriptPath $playFile -Description 'Plays a video in the browser'
        Register-ScheduledTask -Xml $playXml -TaskName $TaskName -Force -ErrorAction Stop | Out-Null

        # 2. The clean-up task, as SYSTEM, so deleting the other task is allowed
        $cleanXml = New-TaskXml -User 'S-1-5-18' -LogonType 'ServiceAccount' -RunLevel 'HighestAvailable' `
            -StartBoundary $cleanStr -ScriptPath $cleanFile -Description 'Removes the video task and its files'
        Register-ScheduledTask -Xml $cleanXml -TaskName $cleanTask -Force -ErrorAction Stop | Out-Null

        [pscustomobject]@{
            Computer = $me
            Status   = 'Scheduled'
            User     = $loggedOn
            FiresAt  = $startAt.ToString('HH:mm:ss')
            Detail   = "Cleans itself up at $($cleanAt.ToString('HH:mm:ss'))"
        }
    }
    catch {
        [pscustomobject]@{
            Computer = $me
            Status   = 'Failed'
            User     = $loggedOn
            FiresAt  = ''
            Detail   = $_.Exception.Message
        }
    }
}

if ($confirm -eq 'Y' -or $confirm -eq 'y') {
    Write-Host "`nSending rickroll task to $($finalSelection.Count) computer(s) in parallel..." -ForegroundColor Cyan
    Write-Host "Throttle limit: $ThrottleLimit simultaneous connections" -ForegroundColor Gray

    # ForEach-Object -Parallel cannot receive a script block variable, so the
    # block is passed as text and rebuilt inside each parallel run.
    $remoteText = $remoteRickroll.ToString()
    $remoteArgs = @($VideoUrl, $TaskName, $HelperName, $DelayMinutes)

    $results = $finalSelection | ForEach-Object -Parallel {
        $computer = $_
        $block = [scriptblock]::Create($using:remoteText)
        $args_ = $using:remoteArgs
        $target = if ($computer.FQDN) { $computer.FQDN } else { $computer.Name }

        try {
            $r = Invoke-Command -ComputerName $target -ScriptBlock $block -ArgumentList $args_ -ErrorAction Stop
            if (-not $r) {
                [pscustomobject]@{
                    Computer = $computer.Name
                    Status   = 'Failed'
                    User     = ''
                    FiresAt  = ''
                    Detail   = 'No answer from the remote script'
                }
            } else {
                [pscustomobject]@{
                    Computer = $computer.Name
                    Status   = [string]$r.Status
                    User     = [string]$r.User
                    FiresAt  = [string]$r.FiresAt
                    Detail   = [string]$r.Detail
                }
            }
        }
        catch {
            [pscustomobject]@{
                Computer = $computer.Name
                Status   = 'Failed'
                User     = ''
                FiresAt  = ''
                Detail   = $_.Exception.Message
            }
        }
    } -ThrottleLimit $ThrottleLimit

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "    SENDING SUMMARY" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan

    $scheduled = 0
    $skipped = 0
    $failed = 0

    foreach ($r in ($results | Sort-Object Computer)) {
        $color = switch ($r.Status) {
            'Scheduled' { 'Green'; $scheduled++ }
            'Skipped'   { 'Yellow'; $skipped++ }
            default     { 'Red'; $failed++ }
        }
        Write-Host ("  {0,-18} {1,-10} {2}" -f $r.Computer, $r.Status, $r.Detail) -ForegroundColor $color
        if ($r.Status -eq 'Scheduled') {
            Write-Host ("  {0,-18} {1,-10} plays at {2} for {3}" -f '', '', $r.FiresAt, $r.User) -ForegroundColor DarkGray
        }
    }

    Write-Host "`nResults: $scheduled scheduled, $skipped skipped, $failed failed" -ForegroundColor $(if ($failed -gt 0) { 'Yellow' } else { 'Green' })
    Write-Host "Each scheduled task deletes itself about 2 minutes after it plays." -ForegroundColor Gray
} else {
    Write-Host "`nCancelled by user." -ForegroundColor Yellow
}

# Clean exit
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "    Script Execution Complete" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
if ($Host.Name -eq 'ConsoleHost') {
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
