# Remote_Shutdown.ps1
# Remote Computer Shutdown Manager
# Run from Domain Controller with Administrator privileges
# Requires: ActiveDirectory module, PowerShell 7.0+

#Requires -Version 7.0

Import-Module ActiveDirectory -ErrorAction Stop

$domainName = (Get-ADDomain).Name

function Show-Menu {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "   REMOTE COMPUTER SHUTDOWN MANAGER" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Domain: $domainName" -ForegroundColor Gray
    Write-Host ""
    Write-Host "1. Shutdown Remote Computers"
    Write-Host "2. Restart Remote Computers"
    Write-Host "3. Log Off Users on Remote Computers"
    Write-Host "4. Abort Scheduled Shutdown"
    Write-Host "5. Exit"
    Write-Host ""
}

function Select-Computers {
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
        Read-Host "`nPress Enter to continue"
        return @()
    }
    
    # Display OUs in Out-GridView for selection
    Write-Host "Please select an OU from the popup window..." -ForegroundColor Yellow
    $selectedOU = $allOUs | Select-Object Name, @{Name="Computers";Expression={$_.ComputerCount}}, DistinguishedName | 
        Out-GridView -Title "Select Target OU (Use Filter to Search)" -OutputMode Single
    
    if (-not $selectedOU) {
        Write-Host "No OU selected." -ForegroundColor Red
        return @()
    }
    
    Write-Host "`nSelected OU: $($selectedOU.Name)" -ForegroundColor Green
    Write-Host "Scanning for online computers..." -ForegroundColor Cyan
    
    # Get enabled computers from selected OU
    $computers = Get-ADComputer -Filter {Enabled -eq $true} -SearchBase $selectedOU.DistinguishedName -Properties DNSHostName, OperatingSystem -ErrorAction SilentlyContinue | 
        Sort-Object Name
    
    if ($computers.Count -eq 0) {
        Write-Host "No enabled computers found in this OU." -ForegroundColor Red
        Read-Host "`nPress Enter to continue"
        return @()
    }
    
    Write-Host "Testing connectivity for $($computers.Count) computers..." -ForegroundColor Cyan
    Write-Host "This may take a moment..." -ForegroundColor Gray
    
    # Test online status using parallel processing
    $computerList = @()
    foreach ($computer in $computers) {
        $computerList += [PSCustomObject]@{
            Name = $computer.Name
            DNSHostName = if ($computer.DNSHostName) { $computer.DNSHostName } else { $computer.Name }
            OperatingSystem = $computer.OperatingSystem
        }
    }
    
    # Test connections in parallel
    $results = $computerList | ForEach-Object -Parallel {
        $comp = $_
        $isOnline = Test-Connection -ComputerName $comp.DNSHostName -Count 1 -Quiet -ErrorAction SilentlyContinue
        
        return [PSCustomObject]@{
            Name = $comp.Name
            FQDN = $comp.DNSHostName
            OperatingSystem = $comp.OperatingSystem
            IsOnline = $isOnline
        }
    } -ThrottleLimit 20
    
    $onlineComputers = $results | Where-Object { $_.IsOnline } | ForEach-Object {
        [PSCustomObject]@{
            Name = $_.Name
            FQDN = $_.FQDN
            OperatingSystem = $_.OperatingSystem
        }
    }
    $offlineCount = ($results | Where-Object { -not $_.IsOnline }).Count
    
    Write-Host "`nResults: $($onlineComputers.Count) online, $offlineCount offline" -ForegroundColor Cyan
    
    if ($onlineComputers.Count -eq 0) {
        Write-Host "No online computers found in this OU." -ForegroundColor Red
        Read-Host "`nPress Enter to continue"
        return @()
    }
    
    # Show online computers in grid view for multi-selection
    Write-Host "`nPlease select computers from the popup window..." -ForegroundColor Yellow
    Write-Host "Use Ctrl+Click for multiple selection" -ForegroundColor Gray
    
    $selectedComputers = $onlineComputers | 
        Select-Object Name, OperatingSystem |
        Out-GridView -Title "Select Target Computers (Ctrl+Click for multiple)" -OutputMode Multiple
    
    if (-not $selectedComputers -or $selectedComputers.Count -eq 0) {
        Write-Host "No computers selected." -ForegroundColor Red
        return @()
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
    
    return $finalSelection
}

function Send-ShutdownCommand($Computers, $Action, $Timeout, $Comment) {
    $actionFlag = switch ($Action) {
        "shutdown" { "/s" }
        "restart" { "/r" }
        "logoff" { "/l" }
    }
    
    Write-Host "`nExecuting $Action on $($Computers.Count) computers in parallel..." -ForegroundColor Cyan
    Write-Host "Throttle limit: 10 simultaneous connections" -ForegroundColor Gray
    
    if ($Action -eq "logoff") {
        # Use CIM for remote logoff (more reliable than shutdown /l /m)
        $results = $Computers | ForEach-Object -Parallel {
            $computer = $_
            try {
                $cimSession = New-CimSession -ComputerName $computer.Name -ErrorAction Stop
                $result = Invoke-CimMethod -ClassName Win32_OperatingSystem `
                    -MethodName Win32Shutdown `
                    -Arguments @{Flags=4} `
                    -CimSession $cimSession
                Remove-CimSession $cimSession
                
                if ($result.ReturnValue -eq 0) {
                    return [PSCustomObject]@{
                        Computer = $computer.Name
                        Success = $true
                        Message = "OK"
                    }
                } else {
                    return [PSCustomObject]@{
                        Computer = $computer.Name
                        Success = $false
                        Message = "Return code: $($result.ReturnValue)"
                    }
                }
            }
            catch {
                return [PSCustomObject]@{
                    Computer = $computer.Name
                    Success = $false
                    Message = $_.Exception.Message
                }
            }
        } -ThrottleLimit 10
    }
    else {
        # Use shutdown.exe for shutdown/restart
        $results = $Computers | ForEach-Object -Parallel {
            $computer = $_
            $flag = $using:actionFlag
            $timeout = $using:Timeout
            $comment = $using:Comment
            
            try {
                $process = Start-Process -FilePath "shutdown.exe" `
                    -ArgumentList "/m \\$($computer.Name) $flag /t $timeout /c `"$comment`" /f" `
                    -Wait -PassThru -NoNewWindow -ErrorAction Stop
                
                if ($process.ExitCode -eq 0) {
                    return [PSCustomObject]@{
                        Computer = $computer.Name
                        Success = $true
                        Message = "OK"
                    }
                } else {
                    return [PSCustomObject]@{
                        Computer = $computer.Name
                        Success = $false
                        Message = "Exit code: $($process.ExitCode)"
                    }
                }
            }
            catch {
                return [PSCustomObject]@{
                    Computer = $computer.Name
                    Success = $false
                    Message = $_.Exception.Message
                }
            }
        } -ThrottleLimit 10
    }
    
    # Process results
    $ok = ($results | Where-Object { $_.Success }).Count
    $fail = ($results | Where-Object { -not $_.Success }).Count
    
    # Display individual results
    Write-Host "`nDetailed Results:" -ForegroundColor Cyan
    foreach ($result in $results) {
        if ($result.Success) {
            Write-Host "  $($result.Computer)... $($result.Message)" -ForegroundColor Green
        } else {
            Write-Host "  $($result.Computer)... $($result.Message)" -ForegroundColor Red
        }
    }
    
    # Display summary
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Summary: $ok succeeded, $fail failed" -ForegroundColor $(if ($fail -eq 0) { "Green" } else { "Yellow" })
    if ($Action -ne "logoff" -and $ok -gt 0) {
        Write-Host "Computers will $Action in $Timeout seconds." -ForegroundColor Yellow
    }
    Write-Host "========================================" -ForegroundColor Cyan
}

function Abort-Shutdown($Computers) {
    Write-Host "`nAborting scheduled shutdowns on $($Computers.Count) computers in parallel..." -ForegroundColor Cyan
    
    $results = $Computers | ForEach-Object -Parallel {
        $computer = $_
        try {
            $process = Start-Process -FilePath "shutdown.exe" `
                -ArgumentList "/m \\$($computer.Name) /a" `
                -Wait -PassThru -NoNewWindow -ErrorAction Stop
            
            if ($process.ExitCode -eq 0) {
                return [PSCustomObject]@{
                    Computer = $computer.Name
                    Success = $true
                    Message = "OK"
                }
            } else {
                return [PSCustomObject]@{
                    Computer = $computer.Name
                    Success = $false
                    Message = "No shutdown in progress"
                }
            }
        }
        catch {
            return [PSCustomObject]@{
                Computer = $computer.Name
                Success = $false
                Message = $_.Exception.Message
            }
        }
    } -ThrottleLimit 10
    
    $ok = ($results | Where-Object { $_.Success }).Count
    $fail = ($results | Where-Object { -not $_.Success }).Count
    
    Write-Host "`nDetailed Results:" -ForegroundColor Cyan
    foreach ($result in $results) {
        if ($result.Success) {
            Write-Host "  $($result.Computer)... $($result.Message)" -ForegroundColor Green
        } else {
            Write-Host "  $($result.Computer)... $($result.Message)" -ForegroundColor Red
        }
    }
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Summary: $ok succeeded, $fail failed" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
}

function Get-TimeoutInput {
    Write-Host "`nEnter shutdown delay in seconds:" -ForegroundColor Cyan
    Write-Host "  0 = Immediately"
    Write-Host "  30 = 30 seconds"
    Write-Host "  60 = 1 minute"
    Write-Host "  300 = 5 minutes"
    Write-Host "  600 = 10 minutes"
    Write-Host "  3600 = 1 hour"
    Write-Host ""
    $timeout = Read-Host "Delay (seconds, default 60)"
    
    if (-not $timeout -or $timeout -eq "") { $timeout = 60 }
    return [int]$timeout
}

# Main loop
do {
    Show-Menu
    $choice = Read-Host "Select option (1-5)"
    
    switch ($choice) {
        "1" {
            $computers = Select-Computers
            if ($computers -and $computers.Count -gt 0) {
                $timeout = Get-TimeoutInput
                $comment = Read-Host "Comment/reason for shutdown"
                if (-not $comment) { $comment = "Administrator-initiated shutdown" }
                
                Write-Host "`nWARNING: This will SHUTDOWN $($computers.Count) computer(s) in $timeout seconds!" -ForegroundColor Red
                Write-Host "Affected computers:" -ForegroundColor Yellow
                foreach ($c in $computers) { Write-Host "  - $($c.Name)" -ForegroundColor Gray }
                
                $confirm = Read-Host "`nType 'YES' to confirm"
                
                if ($confirm -eq "YES") {
                    Send-ShutdownCommand -Computers $computers -Action "shutdown" -Timeout $timeout -Comment $comment
                }
                else { Write-Host "Cancelled." -ForegroundColor Gray }
            }
            Read-Host "`nPress Enter to continue"
        }
        "2" {
            $computers = Select-Computers
            if ($computers -and $computers.Count -gt 0) {
                $timeout = Get-TimeoutInput
                $comment = Read-Host "Comment/reason for restart"
                if (-not $comment) { $comment = "Administrator-initiated restart" }
                
                Write-Host "`nWARNING: This will RESTART $($computers.Count) computer(s) in $timeout seconds!" -ForegroundColor Red
                Write-Host "Affected computers:" -ForegroundColor Yellow
                foreach ($c in $computers) { Write-Host "  - $($c.Name)" -ForegroundColor Gray }
                
                $confirm = Read-Host "`nType 'YES' to confirm"
                
                if ($confirm -eq "YES") {
                    Send-ShutdownCommand -Computers $computers -Action "restart" -Timeout $timeout -Comment $comment
                }
                else { Write-Host "Cancelled." -ForegroundColor Gray }
            }
            Read-Host "`nPress Enter to continue"
        }
        "3" {
            $computers = Select-Computers
            if ($computers -and $computers.Count -gt 0) {
                Write-Host "`nWARNING: This will LOG OFF users on $($computers.Count) computer(s)!" -ForegroundColor Red
                Write-Host "Affected computers:" -ForegroundColor Yellow
                foreach ($c in $computers) { Write-Host "  - $($c.Name)" -ForegroundColor Gray }
                
                $confirm = Read-Host "`nType 'YES' to confirm"
                
                if ($confirm -eq "YES") {
                    Send-ShutdownCommand -Computers $computers -Action "logoff" -Timeout 0 -Comment "Administrator-initiated logoff"
                }
                else { Write-Host "Cancelled." -ForegroundColor Gray }
            }
            Read-Host "`nPress Enter to continue"
        }
        "4" {
            $computers = Select-Computers
            if ($computers -and $computers.Count -gt 0) {
                Write-Host "`nAborting scheduled shutdowns on:" -ForegroundColor Cyan
                foreach ($c in $computers) { Write-Host "  - $($c.Name)" -ForegroundColor Gray }
                
                $confirm = Read-Host "`nContinue? (y/n)"
                
                if ($confirm -eq "y") {
                    Abort-Shutdown -Computers $computers
                }
                else { Write-Host "Cancelled." -ForegroundColor Gray }
            }
            Read-Host "`nPress Enter to continue"
        }
        "5" {
            Write-Host "Exiting..." -ForegroundColor Gray
        }
        default {
            Write-Host "Invalid option" -ForegroundColor Red
            Start-Sleep 1
        }
    }
} while ($choice -ne "5")

# Clean exit
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "    Script Execution Complete" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
if ($Host.Name -eq 'ConsoleHost') {
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
Exit 0