# Install-DFS.ps1
# Installs DFS Namespace and all required dependencies

param(
    [switch]$InstallManagementTools = $true,
    [switch]$InstallReplication = $false,
    [switch]$RestartIfNeeded = $false,
    [switch]$WhatIf
)

function Write-Status {
    param(
        [string]$Message, 
        [string]$Status, 
        [string]$Color = "White"
    )
    
    $symbol = switch ($Status) {
        "Success" { "[OK]" }
        "Error"   { "[ERROR]" }
        "Warning" { "[WARN]" }
        "Info"    { "[INFO]" }
        default   { "[  ]" }
    }
    
    Write-Host "$symbol $Message" -ForegroundColor $Color
}

function Test-Admin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Install-DFSFeatures {
    Write-Host "`n" + ("="*60) -ForegroundColor Cyan
    Write-Host "INSTALLING DFS FEATURES" -ForegroundColor Yellow
    Write-Host ("="*60) -ForegroundColor Cyan
    
    # Core DFS features
    $features = @(
        @{Name="FS-DFS-Namespace"; Description="DFS Namespace"}
        @{Name="FS-DFS-Replication"; Description="DFS Replication"; Optional=$true}
    )
    
    foreach ($feature in $features) {
        if ($feature.Optional -and (-not $InstallReplication)) {
            continue
        }
        
        Write-Status -Message "Checking: $($feature.Description)" -Status "Info" -Color "Cyan"
        
        $featureInfo = Get-WindowsFeature -Name $feature.Name
        
        if ($featureInfo.Installed) {
            Write-Status -Message "Already installed: $($feature.Description)" -Status "Success" -Color "Green"
        } else {
            if ($WhatIf) {
                Write-Status -Message "[WhatIf] Would install: $($feature.Description)" -Status "Info" -Color "DarkGray"
            } else {
                try {
                    Write-Host "  Installing..." -ForegroundColor Gray
                    $result = Install-WindowsFeature -Name $feature.Name -IncludeManagementTools
                    
                    if ($result.Success) {
                        Write-Status -Message "Installed: $($feature.Description)" -Status "Success" -Color "Green"
                    } else {
                        Write-Status -Message "Install may require restart" -Status "Warning" -Color "Yellow"
                    }
                } catch {
                    Write-Status -Message "Failed to install: $_" -Status "Error" -Color "Red"
                }
            }
        }
    }
}

function Install-ManagementTools {
    if (-not $InstallManagementTools) {
        return
    }
    
    Write-Host "`n" + ("="*60) -ForegroundColor Cyan
    Write-Host "INSTALLING MANAGEMENT TOOLS" -ForegroundColor Yellow
    Write-Host ("="*60) -ForegroundColor Cyan
    
    $tools = @(
        @{Name="RSAT-DFS-Mgmt-Con"; Description="DFS Management Tools"}
        @{Name="RSAT-AD-PowerShell"; Description="Active Directory PowerShell"}
        @{Name="RSAT-AD-Tools"; Description="Active Directory Tools"}
    )
    
    foreach ($tool in $tools) {
        Write-Status -Message "Checking: $($tool.Description)" -Status "Info" -Color "Cyan"
        
        $toolInfo = Get-WindowsFeature -Name $tool.Name -ErrorAction SilentlyContinue
        
        if ($toolInfo -and $toolInfo.Installed) {
            Write-Status -Message "Already installed: $($tool.Description)" -Status "Success" -Color "Green"
        } else {
            if ($WhatIf) {
                Write-Status -Message "[WhatIf] Would install: $($tool.Description)" -Status "Info" -Color "DarkGray"
            } else {
                try {
                    Write-Host "  Installing..." -ForegroundColor Gray
                    
                    # Different installation methods for different Windows versions
                    if ($PSVersionTable.PSVersion.Major -ge 5) {
                        # Windows Server 2016/2019/2022 or Windows 10/11
                        if ($tool.Name -like "RSAT*") {
                            # Use DISM for RSAT on newer systems
                            Enable-WindowsOptionalFeature -Online -FeatureName $tool.Name -NoRestart
                        } else {
                            Install-WindowsFeature -Name $tool.Name
                        }
                    } else {
                        # Windows Server 2012 R2
                        Install-WindowsFeature -Name $tool.Name
                    }
                    
                    Write-Status -Message "Installed: $($tool.Description)" -Status "Success" -Color "Green"
                } catch {
                    Write-Status -Message "Failed to install: $_" -Status "Error" -Color "Red"
                }
            }
        }
    }
}

function Install-Prerequisites {
    Write-Host "`n" + ("="*60) -ForegroundColor Cyan
    Write-Host "INSTALLING PREREQUISITES" -ForegroundColor Yellow
    Write-Host ("="*60) -ForegroundColor Cyan
    
    $prereqs = @(
        "FileAndStorage-Services",
        "Storage-Services",
        "FS-FileServer"
    )
    
    foreach ($prereq in $prereqs) {
        Write-Status -Message "Checking: $prereq" -Status "Info" -Color "Cyan"
        
        $feature = Get-WindowsFeature -Name $prereq -ErrorAction SilentlyContinue
        
        if ($feature -and $feature.Installed) {
            Write-Status -Message "Already installed: $prereq" -Status "Success" -Color "Green"
        } else {
            if ($WhatIf) {
                Write-Status -Message "[WhatIf] Would install: $prereq" -Status "Info" -Color "DarkGray"
            } else {
                try {
                    Write-Host "  Installing..." -ForegroundColor Gray
                    Install-WindowsFeature -Name $prereq
                    Write-Status -Message "Installed: $prereq" -Status "Success" -Color "Green"
                } catch {
                    Write-Status -Message "Note: $prereq may not be required or already satisfied" -Status "Warning" -Color "Yellow"
                }
            }
        }
    }
}

function Verify-Installation {
    Write-Host "`n" + ("="*60) -ForegroundColor Cyan
    Write-Host "VERIFICATION" -ForegroundColor Yellow
    Write-Host ("="*60) -ForegroundColor Cyan
    
    # Check installed features
    $installedFeatures = @()
    
    $featuresToCheck = @(
        "FS-DFS-Namespace",
        "FS-DFS-Replication"
    )
    
    foreach ($feature in $featuresToCheck) {
        $feat = Get-WindowsFeature -Name $feature
        if ($feat.Installed) {
            Write-Status -Message "$($feat.DisplayName): INSTALLED" -Status "Success" -Color "Green"
            $installedFeatures += $feature
        } else {
            Write-Status -Message "$($feat.DisplayName): NOT INSTALLED" -Status "Error" -Color "Red"
        }
    }
    
    # Check PowerShell module
    try {
        Import-Module DFSN -ErrorAction Stop
        Write-Status -Message "DFS PowerShell module: LOADED" -Status "Success" -Color "Green"
    } catch {
        Write-Status -Message "DFS PowerShell module: NOT AVAILABLE" -Status "Warning" -Color "Yellow"
    }
    
    # Check services
    Write-Host "`nDFS Services:" -ForegroundColor Cyan
    $services = @("DFS", "DFSR", "DfsMgmt")
    foreach ($service in $services) {
        try {
            $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
            if ($svc) {
                $color = if ($svc.Status -eq "Running") { "Green" } else { "Yellow" }
                Write-Host "  $($service.PadRight(10)): $($svc.Status)" -ForegroundColor $color
            } else {
                Write-Host "  $($service.PadRight(10)): Not found" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  $($service.PadRight(10)): Error checking" -ForegroundColor Red
        }
    }
    
    return $installedFeatures
}

function Start-DFSServices {
    Write-Host "`n" + ("="*60) -ForegroundColor Cyan
    Write-Host "STARTING DFS SERVICES" -ForegroundColor Yellow
    Write-Host ("="*60) -ForegroundColor Cyan
    
    $services = @{
        "DFS" = "DFS Namespace"
        "DFSR" = "DFS Replication"
    }
    
    foreach ($service in $services.Keys) {
        try {
            $svc = Get-Service -Name $service -ErrorAction Stop
            if ($svc.Status -ne "Running") {
                if ($WhatIf) {
                    Write-Status -Message "[WhatIf] Would start: $($services[$service])" -Status "Info" -Color "DarkGray"
                } else {
                    Write-Host "Starting $($services[$service])..." -ForegroundColor Gray
                    Start-Service -Name $service
                    Write-Status -Message "Started: $($services[$service])" -Status "Success" -Color "Green"
                }
            } else {
                Write-Status -Message "Already running: $($services[$service])" -Status "Success" -Color "Green"
            }
        } catch {
            Write-Status -Message "Service $service not found or cannot start: $_" -Status "Warning" -Color "Yellow"
        }
    }
}

# Main Execution
Write-Host "DFS INSTALLATION SCRIPT" -ForegroundColor Magenta
Write-Host "Computer: $env:COMPUTERNAME" -ForegroundColor Gray
Write-Host "Date: $(Get-Date)" -ForegroundColor Gray
Write-Host ("-" * 60) -ForegroundColor Gray

# Check if running as administrator
if (-not (Test-Admin)) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Yellow
    exit 1
}

# Check Windows version
$os = Get-WmiObject Win32_OperatingSystem
Write-Host "OS: $($os.Caption)" -ForegroundColor Gray
Write-Host "Version: $($os.Version)" -ForegroundColor Gray

# Check if server
$isServer = $os.ProductType -eq 3  # Server OS
if (-not $isServer) {
    Write-Host "WARNING: This appears to be a client OS. Some features may not be available." -ForegroundColor Yellow
}

# Installation steps
Install-Prerequisites
Install-DFSFeatures
Install-ManagementTools

# Verify installation
$installedFeatures = Verify-Installation

# Start services if features were installed
if ($installedFeatures.Count -gt 0 -and (-not $WhatIf)) {
    Start-DFSServices
}

# Restart if needed
if ($RestartIfNeeded -and (-not $WhatIf)) {
    $restartNeeded = Get-WindowsFeature | Where-Object {$_.InstallState -eq "Pending"} | Select-Object -First 1
    if ($restartNeeded) {
        Write-Host "`n" + ("="*60) -ForegroundColor Cyan
        Write-Host "RESTART REQUIRED" -ForegroundColor Yellow
        Write-Host ("="*60) -ForegroundColor Cyan
        Write-Host "Some features require a system restart to complete installation." -ForegroundColor Yellow
        
        $choice = Read-Host "Restart now? (Y/N)"
        if ($choice -eq 'Y' -or $choice -eq 'y') {
            Restart-Computer -Force
        }
    }
}

Write-Host "`n" + ("="*60) -ForegroundColor Cyan
Write-Host "INSTALLATION COMPLETE" -ForegroundColor Green
Write-Host ("="*60) -ForegroundColor Cyan

Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "1. Create DFS namespace: New-DfsnRoot -Path '\\SOFT.NET\UserRedirects' -TargetPath '\\SERVER-PC\Share'" -ForegroundColor White
Write-Host "2. Import module: Import-Module DFSN" -ForegroundColor White
Write-Host "3. Check installation: Get-DfsnRoot" -ForegroundColor White