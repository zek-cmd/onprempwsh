# Send_Message.ps1
# Remote Message Sender
# Run from Domain Controller with Administrator privileges
# Requires: ActiveDirectory module

Import-Module ActiveDirectory -ErrorAction Stop

$domainName = (Get-ADDomain).Name

Clear-Host
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    REMOTE MESSAGE SENDER" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Domain: $domainName" -ForegroundColor Gray
Write-Host ""

# Get custom message from user
Write-Host "=== MESSAGE CONFIGURATION ===" -ForegroundColor Cyan
do {
    $CustomMessage = Read-Host "Enter the message to send to users"
    if ([string]::IsNullOrWhiteSpace($CustomMessage)) {
        Write-Host "Message cannot be empty. Please enter a message." -ForegroundColor Red
    }
} while ([string]::IsNullOrWhiteSpace($CustomMessage))

Write-Host "`nMessage Preview: '$CustomMessage'" -ForegroundColor Magenta

# Select OU and computers
Write-Host "`n=== SELECT COMPUTERS ===" -ForegroundColor Cyan
Write-Host "Loading OUs..." -ForegroundColor Yellow

# Get all OUs and add Domain Root option
$domainDN = (Get-ADDomain).DistinguishedName

$allOUs = @()

# Add Domain Root as first option
$domainComps = @(Get-ADComputer -Filter {Enabled -eq $true} -SearchBase $domainDN -SearchScope OneLevel -ErrorAction SilentlyContinue)
if ($domainComps.Count -gt 0) {
    $allOUs += [PSCustomObject]@{
        Name = "[Domain Root]"
        DistinguishedName = $domainDN
        ComputerCount = $domainComps.Count
    }
}

# Add all OUs with computers
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

# Display OUs in Out-GridView for selection
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

# Get enabled computers from selected OU
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

# Test online status
$onlineComputers = @()
$offlineCount = 0

foreach ($computer in $computers) {
    $computerName = if ($computer.DNSHostName) { $computer.DNSHostName } else { $computer.Name }
    Write-Host "  Testing $($computer.Name)... " -NoNewline
    
    $isOnline = Test-Connection -ComputerName $computerName -Count 1 -Quiet -ErrorAction SilentlyContinue
    
    if ($isOnline) {
        Write-Host "ONLINE" -ForegroundColor Green
        $onlineComputers += [PSCustomObject]@{
            Name = $computer.Name
            FQDN = $computerName
            OperatingSystem = $computer.OperatingSystem
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

# Show online computers in grid view for multi-selection
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

# Map back to full computer objects
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

# Confirm before sending
Write-Host "`n=== SEND MESSAGE ===" -ForegroundColor Cyan
Write-Host "Message to send: '$CustomMessage'" -ForegroundColor Magenta
Write-Host "Target computers: $($finalSelection.Count)" -ForegroundColor White

$confirm = Read-Host "`nSend message to these computers? (Y/N)"

if ($confirm -eq 'Y' -or $confirm -eq 'y') {
    Write-Host "`nSending messages..." -ForegroundColor Cyan
    
    $SuccessCount = 0
    $FailCount = 0
    
    foreach ($computer in $finalSelection) {
        Write-Host "  Sending to $($computer.Name)... " -NoNewline
        
        try {
            $Job = Invoke-Command -ComputerName $computer.FQDN -ScriptBlock { 
                param($msg)
                msg * $msg 
            } -ArgumentList $CustomMessage -AsJob -ErrorAction Stop
            
            Write-Host "[QUEUED]" -ForegroundColor Green
            $SuccessCount++
        }
        catch {
            Write-Host "[FAILED] - $($_.Exception.Message)" -ForegroundColor Red
            $FailCount++
        }
    }
    
    # Summary
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "    MESSAGE SENDING SUMMARY" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Message: $CustomMessage" -ForegroundColor White
    Write-Host "Target OU: $($selectedOU.Name)" -ForegroundColor White
    Write-Host "Results: $SuccessCount queued, $FailCount failed" -ForegroundColor $(if($FailCount -gt 0){'Yellow'}else{'Green'})
    Write-Host "`nTo check job status: Get-Job | Format-Table -AutoSize" -ForegroundColor Yellow
    Write-Host "To view results: Get-Job | Receive-Job" -ForegroundColor Yellow
} else {
    Write-Host "`nMessage sending cancelled by user." -ForegroundColor Yellow
}

# Clean exit
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "    Script Execution Complete" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
if ($Host.Name -eq 'ConsoleHost') {
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
Exit 0