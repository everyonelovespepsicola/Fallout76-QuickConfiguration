param ()

# 1. Locate the actual project directory by finding the VERSION file (CD out one layer if needed)
$versionFile = $null

if (Test-Path -LiteralPath "VERSION" -PathType Leaf) {
    $versionFile = Get-Item -LiteralPath "VERSION"
}
elseif (Test-Path -LiteralPath "..\VERSION" -PathType Leaf) {
    $versionFile = Get-Item -LiteralPath "..\VERSION"
}
else {
    $versionFile = Get-ChildItem -Path . -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "VERSION" } | Select-Object -First 1
}

if (-not $versionFile) {
    Write-Host "Error: Could not find the 'VERSION' file. Please ensure you are in or near the project folder." -ForegroundColor Red
    exit
}

# Set the working directory to exactly where the VERSION file is located
$scriptDir = $versionFile.DirectoryName
Set-Location $scriptDir

$currentVersion = ([System.IO.File]::ReadAllText($versionFile.FullName)).Trim()
Write-Host "Current version is: $currentVersion" -ForegroundColor Cyan

$newVersion = Read-Host "Enter the new version"

if ([string]::IsNullOrWhiteSpace($newVersion)) {
    Write-Host "No version entered. Exiting." -ForegroundColor Yellow
    exit
}

# 2. Add C# configuration files to the ignore list so we can handle them manually
$ignoreFiles = @("What's new.md", "What's new.rtf", "whatsnew.html", "whatsnewdark.html", "VERSION", "VERSION.cpp", "AssemblyInfo.cs", "Shared.cs", $MyInvocation.MyCommand.Name)
$excludeExtensions = @(".exe", ".dll", ".png", ".jpg", ".jpeg", ".gif", ".ico", ".wav", ".zip", ".7z", ".ba2")

$files = Get-ChildItem -Path $scriptDir -Recurse -File | Where-Object {
    $_.Name -notin $ignoreFiles -and $_.Extension.ToLower() -notin $excludeExtensions
}

$regexCurrent = [regex]::Escape($currentVersion)
$count = 0
$filesToUpdate = @()
$explicitFiles = @("VERSION", "VERSION.cpp", "Fo76ini\Properties\AssemblyInfo.cs", "Fo76ini_Updater\Properties\AssemblyInfo.cs", "Fo76ini\Shared.cs", "Fo76ini_Updater\Shared.cs")
$totalFound = 0

Write-Host "`nScanning files for version '$currentVersion'..." -ForegroundColor Cyan

# Print out our manually targeted files so we can actually see them in the list!
foreach ($file in $explicitFiles) {
    $filePath = Join-Path $scriptDir $file
    if (Test-Path -LiteralPath $filePath -PathType Leaf) {
        Write-Host "Found (Targeted directly): $filePath" -ForegroundColor Yellow
        $totalFound++
    }
}

# Print out files found via Regex
foreach ($file in $files) {
    try {
        $content = [System.IO.File]::ReadAllText($file.FullName)
        if ($content -match $regexCurrent) {
            $filesToUpdate += $file
            Write-Host "Found (By search): $($file.FullName)" -ForegroundColor Yellow
            $totalFound++
        }
    }
    catch {}
}

Write-Host ""
$confirm = Read-Host "Found a total of $totalFound files. Do you want to apply ALL updates? (y/n)"
if ($confirm -notmatch "^y") {
    Write-Host "Operation cancelled." -ForegroundColor Yellow
    exit
}

Write-Host "`nApplying changes..." -ForegroundColor Cyan
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# 3. Update VERSION and VERSION.cpp explicitly using absolute paths
foreach ($fileName in @("VERSION", "VERSION.cpp")) {
    $vFile = Join-Path $scriptDir $fileName
    try {
        [System.IO.File]::WriteAllText($vFile, $newVersion, $utf8NoBom)
        Write-Host "Updated: $vFile" -ForegroundColor Green
        $count++
    }
    catch {
        Write-Host "Failed to update: $vFile - $($_.Exception.Message)" -ForegroundColor Red
    }
}

# 4. Update AssemblyInfo.cs and Shared.cs explicitly using Regex
$csharpFiles = @(
    "Fo76ini\Properties\AssemblyInfo.cs",
    "Fo76ini_Updater\Properties\AssemblyInfo.cs",
    "Fo76ini\Shared.cs",
    "Fo76ini_Updater\Shared.cs"
)

foreach ($file in $csharpFiles) {
    $filePath = Join-Path $scriptDir $file
    if (Test-Path -LiteralPath $filePath -PathType Leaf) {
        try {
            $content = [System.IO.File]::ReadAllText($filePath)
            $content = $content -replace '\[assembly: AssemblyVersion\(".*?"\)\]', "[assembly: AssemblyVersion(`"$newVersion`")]"
            $content = $content -replace '\[assembly: AssemblyFileVersion\(".*?"\)\]', "[assembly: AssemblyFileVersion(`"$newVersion`")]"
            $content = $content -replace 'public const string VERSION = ".*?";', "public const string VERSION = `"$newVersion`";"
            [System.IO.File]::WriteAllText($filePath, $content, $utf8NoBom)
            Write-Host "Updated: $filePath" -ForegroundColor Green
            $count++
        }
        catch {
            Write-Host "Failed to update: $filePath - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# 5. Apply regex replaces to the rest of the scanned files
foreach ($file in $filesToUpdate) {
    try {
        $content = [System.IO.File]::ReadAllText($file.FullName)
        $content = $content -replace $regexCurrent, $newVersion
        [System.IO.File]::WriteAllText($file.FullName, $content, $utf8NoBom)
        Write-Host "Updated: $($file.FullName)" -ForegroundColor Green
        $count++
    }
    catch {
        Write-Host "Skipped or failed to update: $($file.FullName) - $($_.Exception.Message)" -ForegroundColor DarkGray
    }
}

Write-Host "Finished updating $count files to version $newVersion." -ForegroundColor Cyan
