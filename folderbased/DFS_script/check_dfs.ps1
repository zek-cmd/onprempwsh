Get-DfsnRoot | ForEach-Object { 
    Write-Host "`nNamespace: $($_.Path)" -ForegroundColor Cyan; 
    Get-DfsnRootTarget -Path $_.Path | Format-Table -AutoSize 
}