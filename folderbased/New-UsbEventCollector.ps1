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

# How the run-time question looks:
#   $true  = a window listing every domain computer; tick the rows, then press OK
#   $false = a numbered list in the console (type 3,5,8 or 3-7 or all)
# If the window cannot open, the numbered list is used anyway.
$UseGridView      = $true

# Only offer computers from certain OUs. Separate several with a semicolon,
# for example 'CET;ICT12-PERL' or a full OU distinguished name.
#   ''  = the script asks you to pick an OU at run time
#   '-' = the whole domain, no OU question
$OuFilter         = ''

# When an OU is picked, also take the computers in its sub-OUs.
#   $true  = the whole branch (CET plus everything under it)
#   $false = only computers sitting directly in that one OU
$IncludeSubOus    = $true

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

# Is this computer inside that OU? With children, any level below it counts.
function Test-ComputerInOu {
    param([string]$ComputerDn, [string]$OuDn, [bool]$WithChildren)
    if ([string]::IsNullOrWhiteSpace($ComputerDn) -or [string]::IsNullOrWhiteSpace($OuDn)) { return $false }
    if ($WithChildren) {
        return $ComputerDn.EndsWith(',' + $OuDn, 'OrdinalIgnoreCase')
    }
    # Only the OU the computer sits in directly: CN=PC01,OU=CET,...
    $Parts = $ComputerDn.Split([char[]]@(','), 2)
    if ($Parts.Count -lt 2) { return $false }
    return $Parts[1].Equals($OuDn, 'OrdinalIgnoreCase')
}

# OU distinguished name turned into a readable path: SOFTNET\Computer Lab\CET
function Get-OuLabel {
    param([string]$Dn)
    $Names = @($Dn -split ',' | Where-Object { $_ -match '^OU=' } | ForEach-Object { $_ -replace '^OU=', '' })
    if ($Names.Count -eq 0) { return $Dn }
    [array]::Reverse($Names)
    return ($Names -join '\')
}

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
        $DomainComputers = @()
        $AllComputers = @()
        try {
            $AllComputers = @(Get-ADComputer -Filter { Enabled -eq $true } -Property DistinguishedName -ErrorAction Stop |
                Where-Object { $_.Name -ne $env:COMPUTERNAME } |
                Sort-Object Name)
        }
        catch {
            throw "Active Directory could not be read ($($_.Exception.Message)). Put the computer names in the configuration list at the top of this script, separated by commas."
        }
        if ($AllComputers.Count -eq 0) {
            throw "Active Directory returned no enabled computers besides this server - nothing to collect from."
        }

        # --- step 2a: narrow the list to one or more OUs first ---
        $OuDns       = @()
        $WholeDomain = $false
        $NeedOuQuestion = $true

        if ($OuFilter -eq '-') {
            $WholeDomain = $true
            $NeedOuQuestion = $false
            Write-Info "OU filter is '-' - the whole domain will be offered."
        }

        if ($NeedOuQuestion) {
            # Every OU that holds at least one enabled computer
            $DirectOus = @{}
            foreach ($C in $AllComputers) {
                $Dn = [string]$C.DistinguishedName
                if ([string]::IsNullOrWhiteSpace($Dn)) { continue }
                $Parts = $Dn.Split([char[]]@(','), 2)
                if ($Parts.Count -lt 2 -or -not $Parts[1].StartsWith('OU=')) { continue }
                $DirectOus[$Parts[1]] = $true
            }

            $OuRows = @()
            foreach ($OuDn in $DirectOus.Keys) {
                $Count = @($AllComputers | Where-Object {
                        Test-ComputerInOu -ComputerDn ([string]$_.DistinguishedName) -OuDn $OuDn -WithChildren $IncludeSubOus
                    }).Count
                if ($Count -gt 0) {
                    $OuRows += [pscustomobject]@{ OU = (Get-OuLabel $OuDn); Computers = $Count; DN = $OuDn }
                }
            }
            $OuRows = @($OuRows | Sort-Object OU)
            if ($OuRows.Count -eq 0) {
                throw "No enabled computer sits in an OU, so there is nothing to filter by."
            }

            if (-not [string]::IsNullOrWhiteSpace($OuFilter)) {
                # OUs given in the configuration: names or distinguished names
                $Wanted = @($OuFilter -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                foreach ($W in $Wanted) {
                    if ($W -match '^OU=') { $OuDns += $W; continue }
                    $Hit = @($OuRows | Where-Object { $_.OU -eq $W -or $_.OU.EndsWith('\' + $W, 'OrdinalIgnoreCase') })
                    if ($Hit.Count -eq 0) { Write-Warn "No enabled computer found in OU '$W' - left out." }
                    else { $OuDns += @($Hit | ForEach-Object { $_.DN }) }
                }
                $OuDns = @($OuDns | Sort-Object -Unique)
                if ($OuDns.Count -eq 0) {
                    throw "None of the OUs in the configuration (OuFilter = '$OuFilter') hold an enabled computer."
                }
            }
            else {
                Write-Info "Pick the OU(s) to take computers from - their computers are listed right after you pick."
                $OuChosen = @()

                if ($UseGridView -and (Get-Command -Name 'Out-GridView' -ErrorAction SilentlyContinue)) {
                    try {
                        $OuChosen = @($OuRows |
                            Out-GridView -Title 'Tick the OU(s) to collect USB events from (Ctrl+A ticks all), then press OK' -OutputMode Multiple -ErrorAction Stop)
                    }
                    catch {
                        Write-Warn "The OU window could not open - showing the numbered list instead."
                        $OuChosen = @()
                    }
                }

                if ($OuChosen.Count -gt 0) {
                    $OuDns = @($OuChosen | ForEach-Object { $_.DN })
                }
                else {
                    # The list is shown, you pick, and that OU's computers appear at once.
                    while ($true) {
                        Write-Info "Examples: 1   |   1,3   |   all   |   or type an OU name"
                        Write-Info "Blank = every OU (the whole domain)."
                        $Number = 0
                        foreach ($Row in $OuRows) {
                            $Number++
                            Write-Host ("        [{0,2}]  {1,-45} {2,3} computer(s)" -f $Number, $Row.OU, $Row.Computers) -ForegroundColor DarkGray
                        }
                        $Answer = Read-Host "OU number"
                        if ([string]::IsNullOrWhiteSpace($Answer) -or $Answer.Trim() -match '^(all|\*)$') {
                            $WholeDomain = $true
                            break
                        }

                        $Picked = New-Object System.Collections.Generic.List[string]
                        foreach ($Piece in ($Answer -split '[,;]')) {
                            $Piece = $Piece.Trim()
                            if (-not $Piece) { continue }
                            $Num = 0
                            if ([int]::TryParse($Piece, [ref]$Num) -and $Num -ge 1 -and $Num -le $OuRows.Count) {
                                $Picked.Add($OuRows[$Num - 1].DN)
                            }
                            else {
                                $Hit = @($OuRows | Where-Object { $_.OU -eq $Piece -or $_.OU.EndsWith('\' + $Piece, 'OrdinalIgnoreCase') })
                                if ($Hit.Count -gt 0) { foreach ($H in $Hit) { $Picked.Add($H.DN) } }
                                else { Write-Warn "'$Piece' is not in the list - left out." }
                            }
                        }
                        $Picked = @($Picked | Sort-Object -Unique)
                        if ($Picked.Count -eq 0) {
                            Write-Warn "That was not an OU in the list. Try again, or type 'all' for the whole domain."
                            continue
                        }

                        # --- show the computers sitting in the OU(s) just picked ---
                        $OuDns = $Picked
                        $InPick = @($AllComputers | Where-Object {
                                $Dn = [string]$_.DistinguishedName
                                $Inside = $false
                                foreach ($O in $OuDns) {
                                    if (Test-ComputerInOu -ComputerDn $Dn -OuDn $O -WithChildren $IncludeSubOus) { $Inside = $true; break }
                                }
                                $Inside
                            })
                        Write-Ok "$($InPick.Count) computer(s) inside the OU(s) you picked:"
                        foreach ($C in $InPick) {
                            Write-Host ("            {0,-22} {1}" -f $C.Name, $C.DistinguishedName) -ForegroundColor DarkGray
                        }
                        if ($InPick.Count -eq 0) {
                            Write-Warn "That OU holds no enabled computer apart from this server - nothing to collect from."
                            $OuDns = @()
                            continue
                        }
                        if (Read-YesNo "Collect only from these $($InPick.Count) computer(s)?") { break }
                        Write-Info "Fine - pick again from the OU list above."
                        $OuDns = @()
                        $WholeDomain = $false
                    }
                }
            }
        }

        # --- keep only the computers inside the chosen OU(s) ---
        if ($WholeDomain) {
            $DomainComputers = @($AllComputers)
            Write-Info "Whole domain: $($DomainComputers.Count) enabled computer(s)."
        }
        else {
            if ($IncludeSubOus) { Write-Info "The chosen OU(s) include their sub-OUs." }
            else { Write-Info "Only computers sitting directly in the chosen OU(s)." }
            foreach ($O in $OuDns) { Write-Info "OU: $(Get-OuLabel $O)" }
            $DomainComputers = @($AllComputers | Where-Object {
                    $Dn = [string]$_.DistinguishedName
                    $Inside = $false
                    foreach ($O in $OuDns) {
                        if (Test-ComputerInOu -ComputerDn $Dn -OuDn $O -WithChildren $IncludeSubOus) { $Inside = $true; break }
                    }
                    $Inside
                })
            if ($DomainComputers.Count -eq 0) {
                throw "The chosen OU(s) hold no enabled computer besides this server - nothing to collect from."
            }
            Write-Info "In the chosen OU(s): $($DomainComputers.Count) computer(s)."
        }

        $Chosen = @()

        # --- first choice: a window with one row per computer ---
        if ($UseGridView -and (Get-Command -Name 'Out-GridView' -ErrorAction SilentlyContinue)) {
            try {
                $Chosen = @($DomainComputers |
                    Select-Object @{ n = 'Computer'; e = { $_.Name } },
                                  @{ n = 'DNS name'; e = { $_.DNSHostName } },
                                  @{ n = 'Container'; e = { $_.DistinguishedName } } |
                    Out-GridView -Title 'Tick the computers to collect USB events from (Ctrl+A ticks all), then press OK' -OutputMode Multiple -ErrorAction Stop)
            }
            catch {
                Write-Warn "The picker window could not open - showing the numbered list instead."
                $Chosen = @()
            }
        }

        if ($Chosen.Count -gt 0) {
            $Targets = @($Chosen | ForEach-Object { $_.Computer })
            Write-Info "Ticked in the window: $($Targets.Count) computer(s)."
        }
        else {
            # --- second choice: a numbered list in the console ---
            Write-Info "Pick the computers by number."
            Write-Info "Examples: 3,5,8   |   3-7   |   all   |   or type a computer name"
            Write-Info "Blank = every computer in the OU(s) above ($($DomainComputers.Count) found) - you are asked first."
            $Number = 0
            foreach ($C in $DomainComputers) {
                $Number++
                Write-Host ("        [{0,2}]  {1}" -f $Number, $C.Name) -ForegroundColor DarkGray
            }
            $Answer = Read-Host "Numbers"

            if ($Answer.Trim() -match '^(all|\*)$') {
                $Targets = @($DomainComputers | Select-Object -ExpandProperty Name)
            }
            elseif ([string]::IsNullOrWhiteSpace($Answer)) {
                # Blank no longer means 'everything' on its own - it has to be confirmed.
                if (Read-YesNo "Take all $($DomainComputers.Count) computer(s) in the OU(s) above?") {
                    $Targets = @($DomainComputers | Select-Object -ExpandProperty Name)
                }
                else {
                    throw "No computer was picked, so nothing was collected. Run the script again and pick the computers by number."
                }
            }
            else {
                $Pick = New-Object System.Collections.Generic.List[string]
                foreach ($Piece in ($Answer -split '[,;]')) {
                    $Piece = $Piece.Trim()
                    if (-not $Piece) { continue }

                    $RangeMatch = [regex]::Match($Piece, '^(\d+)\s*-\s*(\d+)$')
                    $SingleMatch = [regex]::Match($Piece, '^(\d+)$')

                    if ($RangeMatch.Success) {
                        $From = [int]$RangeMatch.Groups[1].Value
                        $To   = [int]$RangeMatch.Groups[2].Value
                        if ($From -gt $To) { $Swap = $From; $From = $To; $To = $Swap }
                        for ($i = $From; $i -le $To; $i++) {
                            if ($i -ge 1 -and $i -le $DomainComputers.Count) {
                                $Pick.Add($DomainComputers[$i - 1].Name)
                            }
                            else { Write-Warn "Number $i is not in the list - left out." }
                        }
                    }
                    elseif ($SingleMatch.Success) {
                        $i = [int]$SingleMatch.Groups[1].Value
                        if ($i -ge 1 -and $i -le $DomainComputers.Count) {
                            $Pick.Add($DomainComputers[$i - 1].Name)
                        }
                        else { Write-Warn "Number $i is not in the list - left out." }
                    }
                    else {
                        # Not a number, so treat it as a computer name typed by hand.
                        $Pick.Add($Piece)
                    }
                }
                $Targets = @($Pick)
            }
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
