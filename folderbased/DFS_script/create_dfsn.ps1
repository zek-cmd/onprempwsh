# Create-DFSNamespace.ps1
# Interactive script to create DFS namespace with user input

function New-InteractiveDFSNamespace {
    # Clear screen
    Clear-Host
    
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host "    DFS NAMESPACE CREATION WIZARD" -ForegroundColor Yellow
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Check if running as administrator
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if (-not $isAdmin) {
        Write-Host "ERROR: This script requires Administrator privileges!" -ForegroundColor Red
        Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Yellow
        return
    }
    
    # Import DFS module
    try {
        Import-Module DFSN -ErrorAction Stop
        Write-Host "[+] DFS PowerShell module loaded" -ForegroundColor Green
    } catch {
        Write-Host "[-] DFS module not available. Attempting to install..." -ForegroundColor Yellow
        
        # Try to install DFS management tools
        try {
            Install-WindowsFeature -Name RSAT-DFS-Mgmt-Con -IncludeManagementTools -ErrorAction Stop
            Import-Module DFSN -ErrorAction Stop
            Write-Host "[+] DFS module installed and loaded" -ForegroundColor Green
        } catch {
            Write-Host "[-] Failed to install DFS tools. Please install manually:" -ForegroundColor Red
            Write-Host "    Install-WindowsFeature -Name FS-DFS-Namespace,RSAT-DFS-Mgmt-Con" -ForegroundColor White
            return
        }
    }
    
    # Step 1: Get domain name
    Write-Host ""
    Write-Host "STEP 1: DOMAIN INFORMATION" -ForegroundColor Cyan
    Write-Host "------------------------" -ForegroundColor Gray
    
    # Try to detect current domain
    $currentDomain = $env:USERDNSDOMAIN
    if (-not $currentDomain) {
        $currentDomain = (Get-WmiObject Win32_ComputerSystem).Domain
    }
    
    if ($currentDomain) {
        Write-Host "Current domain detected: $currentDomain" -ForegroundColor Gray
        $useCurrent = Read-Host "Use this domain? (Y/N)"
        
        if ($useCurrent -eq 'Y' -or $useCurrent -eq 'y') {
            $domainName = $currentDomain
        } else {
            $domainName = Read-Host "Enter domain name (e.g., SOFT.NET)"
        }
    } else {
        Write-Host "No domain detected. Running in workgroup mode." -ForegroundColor Yellow
        $domainName = Read-Host "Enter domain name for DFS namespace (or leave blank for standalone)"
    }
    
    # Step 2: Get namespace name
    Write-Host ""
    Write-Host "STEP 2: NAMESPACE NAME" -ForegroundColor Cyan
    Write-Host "---------------------" -ForegroundColor Gray
    
    $namespaceName = Read-Host "Enter namespace name (e.g., SharedData, UserRedirects)"
    
    # Remove any leading/trailing spaces and backslashes
    $namespaceName = $namespaceName.Trim().Trim('\')
    
    # Step 3: Get target server
    Write-Host ""
    Write-Host "STEP 3: TARGET SERVER" -ForegroundColor Cyan
    Write-Host "--------------------" -ForegroundColor Gray
    
    $currentServer = $env:COMPUTERNAME
    Write-Host "Current server: $currentServer" -ForegroundColor Gray
    $useCurrentServer = Read-Host "Use this server as target? (Y/N)"
    
    if ($useCurrentServer -eq 'Y' -or $useCurrentServer -eq 'y') {
        $targetServer = $currentServer
    } else {
        $targetServer = Read-Host "Enter target server name (without \\)"
    }
    
    # Step 4: Get target share
    Write-Host ""
    Write-Host "STEP 4: TARGET SHARE" -ForegroundColor Cyan
    Write-Host "--------------------" -ForegroundColor Gray
    
    $shareName = Read-Host "Enter share name (e.g., SharedData$, UserRedirects$)"
    
    # Remove any leading/trailing spaces and backslashes
    $shareName = $shareName.Trim().Trim('\')
    
    # Step 5: Choose namespace type
    Write-Host ""
    Write-Host "STEP 5: NAMESPACE TYPE" -ForegroundColor Cyan
    Write-Host "----------------------" -ForegroundColor Gray
    
    if ([string]::IsNullOrEmpty($domainName)) {
        Write-Host "No domain specified. Creating Standalone namespace." -ForegroundColor Yellow
        $namespaceType = "Standalone"
    } else {
        Write-Host "Available namespace types:" -ForegroundColor White
        Write-Host "  1) Domain-based (Recommended for Active Directory)" -ForegroundColor Gray
        Write-Host "  2) Standalone (No domain dependency)" -ForegroundColor Gray
        
        $typeChoice = Read-Host "Choose namespace type (1 or 2)"
        
        switch ($typeChoice) {
            "1" { $namespaceType = "DomainV2" }
            "2" { $namespaceType = "Standalone" }
            default {
                Write-Host "Using default: Domain-based" -ForegroundColor Yellow
                $namespaceType = "DomainV2"
            }
        }
    }
    
    # Step 6: Build paths
    Write-Host ""
    Write-Host "STEP 6: CONFIGURATION SUMMARY" -ForegroundColor Cyan
    Write-Host "----------------------------" -ForegroundColor Gray
    
    if ($namespaceType -eq "DomainV2") {
        $namespacePath = "\\$domainName\$namespaceName"
    } else {
        $namespacePath = "\\$targetServer\$namespaceName"
    }
    
    $targetPath = "\\$targetServer\$shareName"
    
    Write-Host "DFS Namespace Path: $namespacePath" -ForegroundColor White
    Write-Host "Target Path: $targetPath" -ForegroundColor White
    Write-Host "Namespace Type: $namespaceType" -ForegroundColor White
    
    # Step 7: Check if share exists
    Write-Host ""
    Write-Host "Checking target share..." -ForegroundColor Cyan
    
    if (Test-Path -Path $targetPath -ErrorAction SilentlyContinue) {
        Write-Host "[+] Target share is accessible" -ForegroundColor Green
    } else {
        Write-Host "[-] WARNING: Target share is not accessible!" -ForegroundColor Red
        Write-Host "    Please ensure the share exists and you have permissions." -ForegroundColor Yellow
        
        $createShare = Read-Host "Create the share now? (Y/N)"
        if ($createShare -eq 'Y' -or $createShare -eq 'y') {
            Create-Share -Server $targetServer -ShareName $shareName
        }
    }
    
    # Step 8: Confirm creation
    Write-Host ""
    $confirm = Read-Host "Create DFS namespace with above settings? (Y/N)"
    
    if ($confirm -ne 'Y' -and $confirm -ne 'y') {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        return
    }
    
    # Step 9: Create the namespace
    Write-Host ""
    Write-Host "Creating DFS namespace..." -ForegroundColor Cyan
    
    try {
        $params = @{
            Path        = $namespacePath
            TargetPath  = $targetPath
            Type        = $namespaceType
            Description = "DFS Namespace created via PowerShell on $(Get-Date -Format 'yyyy-MM-dd')"
            ErrorAction = 'Stop'
        }
        
        # Check if namespace already exists
        $existing = Get-DfsnRoot -Path $namespacePath -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host "[-] Namespace already exists!" -ForegroundColor Yellow
            $overwrite = Read-Host "Remove and recreate? (Y/N)"
            
            if ($overwrite -eq 'Y' -or $overwrite -eq 'y') {
                Remove-DfsnRoot -Path $namespacePath -Force
                Write-Host "[+] Existing namespace removed" -ForegroundColor Green
            } else {
                Write-Host "Operation cancelled." -ForegroundColor Yellow
                return
            }
        }
        
        # Create the namespace
        New-DfsnRoot @params
        
        Write-Host "[+] DFS Namespace created successfully!" -ForegroundColor Green
        Write-Host ""
        
        # Display results
        $newNamespace = Get-DfsnRoot -Path $namespacePath
        Write-Host "Namespace Details:" -ForegroundColor Cyan
        $newNamespace | Format-List Path, Type, State, Description
        
        Write-Host "Target Information:" -ForegroundColor Cyan
        Get-DfsnRootTarget -Path $namespacePath | Format-Table -AutoSize
        
        # Additional configuration
        Write-Host ""
        $configure = Read-Host "Configure advanced settings? (Y/N)"
        if ($configure -eq 'Y' -or $configure -eq 'y') {
            Configure-AdvancedSettings -NamespacePath $namespacePath
        }
        
    } catch {
        Write-Host "[-] Error creating DFS namespace: $_" -ForegroundColor Red
        return
    }
    
    # Step 10: Test the namespace
    Write-Host ""
    Write-Host "Testing namespace access..." -ForegroundColor Cyan
    
    if (Test-Path -Path $namespacePath -ErrorAction SilentlyContinue) {
        Write-Host "[+] Namespace is accessible!" -ForegroundColor Green
    } else {
        Write-Host "[-] Namespace created but not immediately accessible" -ForegroundColor Yellow
        Write-Host "    This is normal. It may take a few minutes to replicate." -ForegroundColor Gray
    }
    
    Write-Host ""
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host "    DFS NAMESPACE CREATION COMPLETE" -ForegroundColor Green
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Namespace: $namespacePath" -ForegroundColor White
    Write-Host "Type: $namespaceType" -ForegroundColor White
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "1. Add more targets: New-DfsnRootTarget -Path '$namespacePath' -TargetPath '\\Server2\$shareName'" -ForegroundColor Gray
    Write-Host "2. Add folders: New-DfsnFolder -Path '$namespacePath\FolderName' -TargetPath '\\Server\Share'" -ForegroundColor Gray
    Write-Host "3. Test from client: Test-Path -Path '$namespacePath'" -ForegroundColor Gray
}

function Create-Share {
    param(
        [string]$Server,
        [string]$ShareName
    )
    
    Write-Host "Creating share on $Server..." -ForegroundColor Cyan
    
    # Remove $ from share name for folder path
    $folderName = $ShareName -replace '\$$', ''
    
    # Create local path
    $localPath = "C:\Shares\$folderName"
    
    try {
        # Create directory
        New-Item -Path $localPath -ItemType Directory -Force -ErrorAction Stop
        Write-Host "[+] Directory created: $localPath" -ForegroundColor Green
        
        # Create share
        New-SmbShare -Name $ShareName -Path $localPath -FullAccess "Everyone" -ErrorAction Stop
        Write-Host "[+] Share created: \\$Server\$ShareName" -ForegroundColor Green
        
    } catch {
        Write-Host "[-] Error creating share: $_" -ForegroundColor Red
        Write-Host "    Please create the share manually and try again." -ForegroundColor Yellow
    }
}

function Configure-AdvancedSettings {
    param([string]$NamespacePath)
    
    Write-Host ""
    Write-Host "Advanced Configuration" -ForegroundColor Cyan
    Write-Host "---------------------" -ForegroundColor Gray
    
    Write-Host "Available options:" -ForegroundColor White
    Write-Host "  1) Enable Site Costing (for multi-site environments)" -ForegroundColor Gray
    Write-Host "  2) Set Time-to-Live (TTL) for referrals" -ForegroundColor Gray
    Write-Host "  3) Enable Access-based Enumeration" -ForegroundColor Gray
    Write-Host "  4) Set all recommended settings" -ForegroundColor Gray
    
    $choice = Read-Host "Select option (1-4) or Enter to skip"
    
    try {
        switch ($choice) {
            "1" {
                Set-DfsnRoot -Path $NamespacePath -EnableSiteCosting $true
                Write-Host "[+] Site costing enabled" -ForegroundColor Green
            }
            "2" {
                $ttl = Read-Host "Enter TTL in seconds (default: 1800)"
                if (-not $ttl) { $ttl = 1800 }
                Set-DfsnRoot -Path $NamespacePath -TimeToLiveSec $ttl
                Write-Host "[+] TTL set to $ttl seconds" -ForegroundColor Green
            }
            "3" {
                Set-DfsnRoot -Path $NamespacePath -EnableAccessBasedEnumeration $true
                Write-Host "[+] Access-based enumeration enabled" -ForegroundColor Green
            }
            "4" {
                Set-DfsnRoot -Path $NamespacePath -EnableSiteCosting $true
                Set-DfsnRoot -Path $NamespacePath -TimeToLiveSec 1800
                Set-DfsnRoot -Path $NamespacePath -EnableInsiteReferrals $true
                Write-Host "[+] All recommended settings applied" -ForegroundColor Green
            }
            default {
                Write-Host "No advanced configuration applied." -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "[-] Error configuring settings: $_" -ForegroundColor Red
    }
}

# Run the interactive function
New-InteractiveDFSNamespace

# Optional: Keep console open
Write-Host ""
Read-Host "Press Enter to exit"