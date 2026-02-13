<#
.SYNOPSIS
    Register SearchOpenExe context menu entries in Windows Explorer.
.DESCRIPTION
    Automatically detects the script location and registers right-click
    context menu entries for folders, folder backgrounds, and files.
    Uses HKCU (no admin required for registration itself).
#>

$ErrorActionPreference = "Stop"

# Detect the path to SearchOpenExe.ps1
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$mainScript = Join-Path $scriptDir "SearchOpenExe.ps1"

if (-not (Test-Path $mainScript)) {
    Write-Host "[ERROR] SearchOpenExe.ps1 not found at: $mainScript" -ForegroundColor Red
    Write-Host "Please place Install.ps1 in the same directory as SearchOpenExe.ps1." -ForegroundColor Yellow
    pause
    exit 1
}

# Escape backslashes for registry command value
$escapedPath = $mainScript.Replace('\', '\\')

Write-Host "=== SearchOpenExe Installer ===" -ForegroundColor Cyan
Write-Host "Script path: $mainScript" -ForegroundColor White
Write-Host ""

# Registry entries to create
$entries = @(
    @{
        Path    = "HKCU:\Software\Classes\Directory\shell\SearchOpenExe"
        Label   = "Search Open Processes"
        Command = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$mainScript`" `"%V`""
        Description = "Folder right-click"
    },
    @{
        Path    = "HKCU:\Software\Classes\Directory\Background\shell\SearchOpenExe"
        Label   = "Search Open Processes"
        Command = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$mainScript`" `"%V`""
        Description = "Folder background right-click"
    },
    @{
        Path    = "HKCU:\Software\Classes\*\shell\SearchOpenExe"
        Label   = "Search Locking Processes"
        Command = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$mainScript`" `"%1`""
        Description = "File right-click"
    }
)

$successCount = 0

foreach ($entry in $entries) {
    try {
        # Create shell key
        if (-not (Test-Path $entry.Path)) {
            New-Item -Path $entry.Path -Force | Out-Null
        }
        Set-ItemProperty -Path $entry.Path -Name "(Default)" -Value $entry.Label
        Set-ItemProperty -Path $entry.Path -Name "Icon" -Value "shell32.dll,22"

        # Create command subkey
        $commandPath = Join-Path $entry.Path "command"
        if (-not (Test-Path $commandPath)) {
            New-Item -Path $commandPath -Force | Out-Null
        }
        Set-ItemProperty -Path $commandPath -Name "(Default)" -Value $entry.Command

        Write-Host "[OK] $($entry.Description)" -ForegroundColor Green
        $successCount++
    } catch {
        Write-Host "[FAIL] $($entry.Description): $_" -ForegroundColor Red
    }
}

Write-Host ""
if ($successCount -eq $entries.Count) {
    Write-Host "Installation complete! ($successCount/$($entries.Count) entries registered)" -ForegroundColor Green
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Cyan
    Write-Host "  - Right-click a folder -> 'Show more options' -> 'Search Open Processes'"
    Write-Host "  - Right-click inside a folder -> 'Show more options' -> 'Search Open Processes'"
    Write-Host "  - Right-click a file -> 'Show more options' -> 'Search Locking Processes'"
} else {
    Write-Host "Installation partially complete ($successCount/$($entries.Count)). Check errors above." -ForegroundColor Yellow
}

Write-Host ""
pause
