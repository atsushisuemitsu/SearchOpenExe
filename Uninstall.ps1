<#
.SYNOPSIS
    Remove SearchOpenExe context menu entries from Windows Explorer.
.DESCRIPTION
    Removes all registry entries created by Install.ps1.
#>

$ErrorActionPreference = "Stop"

Write-Host "=== SearchOpenExe Uninstaller ===" -ForegroundColor Cyan
Write-Host ""

$paths = @(
    @{ Path = "HKCU:\Software\Classes\Directory\shell\SearchOpenExe"; Description = "Folder right-click" },
    @{ Path = "HKCU:\Software\Classes\Directory\Background\shell\SearchOpenExe"; Description = "Folder background right-click" },
    @{ Path = "HKCU:\Software\Classes\*\shell\SearchOpenExe"; Description = "File right-click" }
)

$removedCount = 0

foreach ($item in $paths) {
    if (Test-Path $item.Path) {
        try {
            Remove-Item -Path $item.Path -Recurse -Force
            Write-Host "[OK] Removed: $($item.Description)" -ForegroundColor Green
            $removedCount++
        } catch {
            Write-Host "[FAIL] $($item.Description): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "[SKIP] Not found: $($item.Description)" -ForegroundColor Gray
    }
}

Write-Host ""
if ($removedCount -gt 0) {
    Write-Host "Uninstallation complete. Removed $removedCount entry(ies)." -ForegroundColor Green
} else {
    Write-Host "Nothing to remove. SearchOpenExe was not installed." -ForegroundColor Yellow
}

Write-Host ""
pause
