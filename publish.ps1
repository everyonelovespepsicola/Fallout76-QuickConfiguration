param (
    [switch]$SkipInstall
)

$ErrorActionPreference = 'Stop'

# Check if Fo76ini.exe is running
if (Get-Process -Name "Fo76ini" -ErrorAction SilentlyContinue) {
    Write-Warning "Fo76ini.exe is currently running! Please close the executable to compile."
    exit
}

$versionFile = "VERSION"
if (Test-Path $versionFile) {
    $currentVersion = (Get-Content $versionFile).Trim()
    if ($currentVersion -match '^(.*?\.)(\d+)$') {
        $prefix = $matches[1]
        $lastNumStr = $matches[2]
        $lastNumInt = [int]$lastNumStr
        $lastNumInt++
        $newLastNumStr = $lastNumInt.ToString().PadLeft($lastNumStr.Length, '0')
        $newVersion = "$prefix$newLastNumStr"

        $response = Read-Host "Current version is $currentVersion. Do you want to update to $newVersion? (y/N)"
        if ($response -match "^y") {
            [System.IO.File]::WriteAllText((Join-Path (Get-Location) $versionFile), "$newVersion`n")
            Write-Host "Version updated to $newVersion!" -ForegroundColor Green
        }
    }
}

if (-not $SkipInstall) {
    Write-Host "Installing dependencies using winget..."

    # List of basic Winget package IDs
    $packages = @(
        "Python.Python.3.12",
        "ElectronCommunity.rcedit",
        "7zip.7zip",
        "JRSoftware.InnoSetup",
        "JohnMacFarlane.Pandoc",
        "Microsoft.NuGet"
    )

    foreach ($pkg in $packages) {
        Write-Host "Installing $pkg..."
        winget install --id $pkg --accept-package-agreements --accept-source-agreements --silent --no-upgrade
    }

    Write-Host "Installing Visual Studio Build Tools 2022 (with MSBuild & Managed Desktop workloads)..."
    winget install --id Microsoft.VisualStudio.2022.BuildTools --accept-package-agreements --accept-source-agreements --silent --no-upgrade --override "--wait --quiet --add Microsoft.VisualStudio.Workload.ManagedDesktop --add Microsoft.Component.MSBuild"
}

Write-Host "Refreshing Environment Variables..."
# Refresh PATH to ensure new tools are available
$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")

# Add common installation paths for tools that might not automatically add themselves to PATH
$innoSetupPathMachine = "${env:ProgramFiles(x86)}\Inno Setup 6"
$innoSetupPathUser = "$env:LOCALAPPDATA\Programs\Inno Setup 6"
if (Test-Path $innoSetupPathMachine) {
    $env:Path += ";$innoSetupPathMachine"
} elseif (Test-Path $innoSetupPathUser) {
    $env:Path += ";$innoSetupPathUser"
}

$sevenZipPath = "$env:ProgramFiles\7-Zip"
if (Test-Path $sevenZipPath) {
    $env:Path += ";$sevenZipPath"
}

Write-Host "Installing required Python packages..."
# We wrap this in a try-catch in case python is still not immediately resolving in this session
try {
    py -m pip install -r requirements.txt
} catch {
    Write-Warning "Failed to run 'py -m pip install -r requirements.txt'. You might need to restart your terminal for python to be on your PATH."
}

Write-Host "Running publish commands via pack_tool.py (-r -b -p -s)..."
try {
    py pack_tool.py -r -b -p -s
    Write-Host "Publish completed successfully!" -ForegroundColor Green
} catch {
    Write-Error "Failed to run publish script. Check if dependencies are correctly configured."
}
