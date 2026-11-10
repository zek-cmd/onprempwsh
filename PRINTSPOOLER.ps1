# PRINTSPOOLER.ps1
# Restarts the Print Spooler service straight away - no menu, no questions.
# Run it in PowerShell as Administrator:  .\PRINTSPOOLER.ps1

Restart-Service -Name Spooler -Force

Write-Host "Print Spooler restarted." -ForegroundColor Green
Write-Host ("Status: " + (Get-Service -Name Spooler).Status)
