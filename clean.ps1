param (
    [switch]$CleanPips
)

$ErrorActionPreference = 'SilentlyContinue'

Write-Host "Cleaning build directories..." -ForegroundColor Cyan

$directoriesToRemove = @(
    "Publish",
    "Fo76ini\bin",
    "Fo76ini\obj",
    "Fo76ini_Updater\bin",
    "Fo76ini_Updater\obj",
    "ObjectListView\bin",
    "ObjectListView\obj"
)

foreach ($dir in $directoriesToRemove) {
    if (Test-Path $dir) {
        Write-Host "Removing $dir..."
        Remove-Item -Recurse -Force $dir
    }
}

if ($CleanPips) {
    Write-Host "Uninstalling pip packages to start over..." -ForegroundColor Cyan
    # Restore the pip-tools execution status by uninstalling the packages
    py -m pip uninstall -y colorama pip-tools build click pyproject-hooks wheel
}

Write-Host "Clean complete!" -ForegroundColor Green
