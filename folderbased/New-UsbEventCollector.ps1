#Requires -RunAsAdministrator

# ============================================================================
# USB EVENT COLLECTOR - run this ONCE, on the SERVER
# ============================================================================
# setgpo.ps1 turns USB plug / unplug logging ON for every client. Each client then
# writes the events into its OWN log, and nobody can see the other PCs' logs.
#
# This script builds the one thing that fixes that: a Windows Event Collector
# subscription. The server then pulls the USB events from every client it is told
# about and stores them in one local log, 'ForwardedEvents'.
#
# Order of work:
#   1. run setgpo.ps1 on this server so the clients start logging
#   2. run this script on this server
#   3. gpupdate /force on a client, plug in a flash drive, wait a minute
#   4. look at Event Viewer > Windows Logs > ForwardedEvents on this server
#
# Nothing is deleted by this script. An existing subscription with the same name is
# only replaced after you say yes.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================================
# 1. CONFIGURATION - edit these
# ============================================================================

# Name of the subscription. It is also what you pick in Event Viewer later.
$SubscriptionName = 'USB Device Watch'

# Which clients to collect from. "" = the script asks you at run time.
# You can also paste computer names here, or a path to a text file with one name per line.
$ComputerList     = ''

# Events to collect from each client. Two logs are read:
#   Security                            -> 6416 = device connected / disconnected
#   ...DriverFrameworks-UserMode/Operational -> 2003 / 2100 / 2101 / 2102 = the device
#                                          instance, which carries the serial number
$SecurityEventIds = @(6416)
$UsbDriverEventIds = @(2003, 2100, 2101, 2102)

# Also pull file opens on removable disks (4656 / 4663). Only useful when
# $EnableUsbFileAccess is turned on in setgpo.ps1 - it is very noisy.
$IncludeRemovableFileAccess = $false

# How much of the forwarded log Windows keeps on this server, in bytes (1 GB).
$ForwardedLogMaxSizeBytes = 1073741824

# Seconds between heartbeat events, so a dead client shows up as a gap.
$HeartbeatSeconds = 300

# ============================================================================
# HELPERS
# ============================================================================
function Write-Step { param([string]$Text) Write-Host "`n[+] $Text" -ForegroundColor Cyan }
function Write-Ok   { param([string]$Text) Write-Host "    OK  $Text" -ForegroundColor Green }
function Write-Info { param([string]$Text) Write-Host "        $Text" -ForegroundColor Gray }
function Write-Warn { param([string]$Text) Write-Host "    !!  $Text" -ForegroundColor Yellow }

function Read-YesNo {
    param([string]$Prompt)
    while ($true) {
        $Answer = Read-Host "$Prompt [y/N]"
        if ([string]::IsNullOrWhiteSpace($Answer)) { return $false }
        if ($Answer.Trim() -match '^(y|yes)$') { return $true }
        if ($Answer.Trim() -match '^(n|no)$') { return $false }
        Write-Warn "Please answer y or n."
    }
}

$StartTime = Get-Date

try {
    if (-not [System.Environment]::UserInteractive) {
        throw "This script asks questions, so it must run in an interactive PowerShell window."
    }

    # ========================================================================
    # 1/5  THE COLLECTOR SERVICE
    # ========================================================================
    Write-Step "1/5  Windows Event Collector service on this server"
    $Wec = Get-Service -Name 'Wecsvc' -ErrorAction SilentlyContinue
    if (-not $Wec) {
        throw "The Windows Event Collector service (Wecsvc) is not installed. On Windows Server add the feature: Install-WindowsFeature -Name 'Windows-Event-Collector' (or install it with Server Manager)."
    }
    if ($Wec.StartType -ne 'Automatic') {
        Set-Service -Name 'Wecsvc' -StartupType Automatic
        Write-Ok "Set Wecsvc to start automatically."
    }
    if ($Wec.Status -ne 'Running') {
        Start-Service -Name 'Wecsvc'
        Write-Ok "Started the Wecsvc service."
    }
    else { Write-Ok "Wecsvc is running and set to Automatic." }

    # ========================================================================
    # 2/5  WHICH CLIENTS TO COLLECT FROM
    # ========================================================================
    Write-Step "2/5  Which computers to collect from"
    $Targets = @()
    if (-not [string]::IsNullOrWhiteSpace($ComputerList)) {
        if (Test-Path -LiteralPath $ComputerList) {
            $Targets = @(Get-Content -LiteralPath $ComputerList | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }
        else {
            $Targets = @($ComputerList -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }
        Write-Info "From the configuration list: $($Targets.Count) computer(s)."
    }
    else {
        $Suggestion = ''
        try {
            $Suggestion = @(Get-ADComputer -Filter { Enabled -eq $true } -ErrorAction Stop |
                Where-Object { $_.Name -ne $env:COMPUTERNAME } |
                Select-Object -ExpandProperty Name) -join ','
        }
        catch { $Suggestion = '' }

        Write-Info "Type the computer names, separated by commas."
        if ($Suggestion) {
            Write-Info "Blank = every enabled computer in the domain except this server ($((@($Suggestion -split ',')).Count) found)."
        }
        $Answer = Read-Host "Computer names"
        if ([string]::IsNullOrWhiteSpace($Answer)) {
            if (-not $Suggestion) { throw "No computer names were given and Active Directory could not be read. Put the names in the prompt, one list separated by commas." }
            $Targets = @($Suggestion -split ',')
        }
        else {
            $Targets = @($Answer -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }
    }
    $Targets = @($Targets | Sort-Object -Unique)
    if ($Targets.Count -eq 0) { throw "No computers to collect from - nothing was created." }
    Write-Ok "Collecting from $($Targets.Count) computer(s)."
    Write-Info (($Targets | Select-Object -First 8) -join ', ' + $(if ($Targets.Count -gt 8) { ' ...' } else { '' }))

    # ========================================================================
    # 3/5  THE QUERY AND THE SUBSCRIPTION FILE
    # ========================================================================
    Write-Step "3/5  The USB event query"
    $SecSelector = '*[System[(' + (($SecurityEventIds | ForEach-Object { "EventID=$_" }) -join ' or ') + ')]]'
    $DrvSelector = '*[System[(' + (($UsbDriverEventIds | ForEach-Object { "EventID=$_" }) -join ' or ') + ')]]'

    $Queries = New-Object System.Collections.ArrayList
    [void]$Queries.Add([pscustomobject]@{ Path = 'Security'; Select = $SecSelector; What = 'device plugged in / unplugged (6416)' })
    [void]$Queries.Add([pscustomobject]@{ Path = 'Microsoft-Windows-DriverFrameworks-UserMode/Operational'; Select = $DrvSelector; What = 'the USB device itself, with serial number' })
    if ($IncludeRemovableFileAccess) {
        [void]$Queries.Add([pscustomobject]@{ Path = 'Security'; Select = '*[System[(EventID=4656 or EventID=4663)]]'; What = 'files opened on a removable disk' })
    }
    foreach ($Q in $Queries) { Write-Info ("{0}  <- {1}" -f $Q.Path, $Q.What) }

    $QueryList = ($Queries | ForEach-Object {
        '<Query Id="0" Path="{0}"><Select Path="{0}">{1}</Select></Query>' -f $Q.Path, $Q.Select
    }) -join ''
    $QueryXml = '<QueryList>' + $QueryList + '</QueryList>'

    # wecutil takes a subscription XML with the computer list and the query inside it.
    $ComputersXml = ($Targets | ForEach-Object { '<Computer>{0}</Computer>' -f $_ }) -join ''
    $SubscriptionXml = @"
<?xml version="1.0" encoding="UTF-8"?>
<Subscription xmlns="http://schemas.microsoft.com/2006/03/windows/events/subscription">
    <SubscriptionId>$SubscriptionName</SubscriptionId>
    <Uri>http://schemas.microsoft.com/wbem/wsman/1/windows/EventLog</Uri>
    <ConfigurationMode>Custom</ConfigurationMode>
    <Delivery Mode="Push">
        <Batching>
            <MaxItems>40</MaxItems>
            <MaxLatencyTime>900000</MaxLatencyTime>
        </Batching>
        <Heartbeat>
            <Interval>$($HeartbeatSeconds * 1000)</Interval>
        </Heartbeat>
    </Delivery>
    <Query>$([System.Security.SecurityElement]::Escape($QueryXml))</Query>
    <ReadExistingEvents>false</ReadExistingEvents>
    <TransportName>HTTP</TransportName>
    <ContentFormat>RenderedText</ContentFormat>
    <Locale Language="en-US"/>
    <LogFile>ForwardedEvents</LogFile>
    <AllowedSourceNonDomainComputers/>
    <AllowedSourceDomainComputers>O:NSG:BAD:P(A;;GA;;;DC)(A;;GA;;;NS)S:P(AU;FA;GA;;;WD)(AU;SA;GXGW;;;WD)</AllowedSourceDomainComputers>
    <ComputerList>$ComputersXml</ComputerList>
</Subscription>
"@
    $XmlPath = Join-Path $env:TEMP 'zz-usb-subscription.xml'
    [System.IO.File]::WriteAllText($XmlPath, $SubscriptionXml, (New-Object System.Text.UTF8Encoding($false)))
    Write-Ok "Subscription file written: $XmlPath"

    # ========================================================================
    # 4/5  MAKING THE SUBSCRIPTION
    # ========================================================================
    Write-Step "4/5  The subscription '$SubscriptionName'"
    $Existing = @(& wecutil.exe es 2>&1)
    $HasIt = @($Existing | Where-Object { $_.Trim() -eq $SubscriptionName }).Count -gt 0
    if ($HasIt) {
        Write-Warn "A subscription named '$SubscriptionName' already exists."
        if (-not (Read-YesNo -Prompt "Replace it with the new computer list and query?")) {
            Write-Info "Left the existing subscription alone."
            $HasIt = $true
            $Replaced = $false
        }
        else {
            # wecutil has no quiet switch for delete: the name alone is the whole call.
            & wecutil.exe ds $SubscriptionName 2>&1 | Out-Null
            Write-Ok "Old subscription removed."
            $Replaced = $true
        }
    }
    else { $Replaced = $false }

    if (-not $HasIt -or $Replaced) {
        $CreateOut = & wecutil.exe cs $XmlPath 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "Windows would not create the subscription: $($CreateOut -join ' ')"
        }
        Write-Ok "Subscription created on this server."
    }

    $Status = @(& wecutil.exe gr $SubscriptionName 2>&1)
    $StatusText = ($Status -join "`n")
    if ($StatusText -match 'Enabled:\s*true|enabled:\s*true') { Write-Ok "The subscription is enabled." }
    else { Write-Warn "The subscription exists but does not say it is enabled - check it in Event Viewer." }

    # ========================================================================
    # 5/5  ROOM FOR THE EVENTS, AND WHO MAY READ THE CLIENTS' LOGS
    # ========================================================================
    Write-Step "5/5  Storage and permissions"
    $Forwarded = Get-WinEvent -ListLog 'ForwardedEvents' -ErrorAction SilentlyContinue
    if ($Forwarded) {
        if ($Forwarded.MaximumSizeInBytes -ne $ForwardedLogMaxSizeBytes) {
            & wevtutil.exe sl ForwardedEvents /ms:$ForwardedLogMaxSizeBytes 2>&1 | Out-Null
            Write-Ok "ForwardedEvents log size set to $([int]($ForwardedLogMaxSizeBytes / 1MB)) MB."
        }
        else { Write-Ok "ForwardedEvents log is already $([int]($ForwardedLogMaxSizeBytes / 1MB)) MB." }
    }
    else { Write-Warn "The 'ForwardedEvents' log was not found - forwarded events will have nowhere to land." }

    Write-Info "The collector reads each client's Security log as this server's computer account."
    Write-Info "On every client that account must be allowed to read the log (group 'Event Log Readers')."
    Write-Info "If no events arrive, add it there - ask the setgpo.ps1 run to cover it, or do it once by hand."

    # ========================================================================
    # DONE
    # ========================================================================
    Write-Host "`n=== Finished in $([int]((Get-Date) - $StartTime).TotalSeconds)s ===" -ForegroundColor Green
    Write-Host "Subscription  : $SubscriptionName"
    Write-Host "Collecting    : $($Targets.Count) computer(s)"
    Write-Host "Stored in     : ForwardedEvents (on $env:COMPUTERNAME)"
    Write-Host "`nHow to use it:" -ForegroundColor Cyan
    Write-Info "1. On a client: 'gpupdate /force', then plug a flash drive in and pull it out."
    Write-Info "2. Wait about a minute, then open Event Viewer on this server."
    Write-Info "3. Windows Logs > ForwardedEvents - the plug and unplug lines appear here, tagged with the computer they came from."
    Write-Info "4. Filter for Event ID 6416 (plug / unplug) and 2003 or 2100 (the device, with its serial number)."
    Write-Info "5. Runtime health of the subscription: 'wecutil gr $SubscriptionName' lists each PC and when it last reported."
    Write-Info "6. Read the events from the console instead:"
    Write-Info "   Get-WinEvent -LogName ForwardedEvents | Where-Object { @(6416,2003,2100,2101,2102) -contains `$_.Id }"
    Write-Host ""
}
catch {
    Write-Host "`n=== Collector setup FAILED ===" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "At line: $($_.InvocationInfo.ScriptLineNumber)  ->  $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkYellow
    exit 1
}
