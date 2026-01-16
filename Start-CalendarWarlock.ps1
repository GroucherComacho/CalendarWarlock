<#
.SYNOPSIS
    Launcher script for CalendarWarlock
.DESCRIPTION
    Checks prerequisites and launches the CalendarWarlock GUI application.
.EXAMPLE
    .\Start-CalendarWarlock.ps1
#>

$ErrorActionPreference = "Stop"
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$mainScript = Join-Path $scriptPath "src\CalendarWarlock.ps1"

# Check PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Error "CalendarWarlock requires PowerShell 5.1 or higher. Current version: $($PSVersionTable.PSVersion)"
    exit 1
}

# Check for required modules
$requiredModules = @(
    @{ Name = "ExchangeOnlineManagement"; MinVersion = "3.0.0" },
    @{ Name = "Microsoft.Graph.Users"; MinVersion = "2.0.0" }
)

$missingModules = @()

foreach ($module in $requiredModules) {
    $installed = Get-Module -ListAvailable -Name $module.Name | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $installed) {
        $missingModules += $module.Name
    }
}

if ($missingModules.Count -gt 0) {
    Write-Host "Missing required PowerShell modules:" -ForegroundColor Yellow
    foreach ($module in $missingModules) {
        Write-Host "  - $module" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "Install missing modules with:" -ForegroundColor Cyan
    Write-Host "  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor White
    Write-Host "  Install-Module -Name Microsoft.Graph -Scope CurrentUser" -ForegroundColor White
    Write-Host ""

    $continue = Read-Host "Continue anyway? (y/N)"
    if ($continue -ne 'y' -and $continue -ne 'Y') {
        exit 1
    }
}

# Launch the main application
Write-Host "Starting CalendarWarlock..." -ForegroundColor Green
& $mainScript
