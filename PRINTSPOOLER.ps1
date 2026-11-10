# PRINTSPOOLER.ps1
# Clears the print queue and restarts the Print Spooler straight away - no menu, no questions.
# Run it in PowerShell as Administrator:  .\PRINTSPOOLER.ps1

$spoolFolder = Join-Path $env:SystemRoot 'System32\spool\PRINTERS'

Stop-Service -Name Spooler -Force

Get-ChildItem -LiteralPath $spoolFolder -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue

Start-Service -Name Spooler

Write-Host "Print queue cleared and Print Spooler restarted." -ForegroundColor Green
Write-Host ("Status: " + (Get-Service -Name Spooler).Status)

Write-Host ""
Write-Host "Press any key to exit..."
try { $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown') } catch { $null = Read-Host }
