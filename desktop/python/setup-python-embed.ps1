# Embedded Python runtime setup script.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File python\setup-python-embed.ps1

$ErrorActionPreference = "Stop"

$pythonVersion = "3.11.9"
$pythonDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($pythonDir)) {
    $scriptPath = $MyInvocation.MyCommand.Path
    if (-not [string]::IsNullOrWhiteSpace($scriptPath)) {
        $pythonDir = Split-Path -Parent $scriptPath
    } else {
        $pythonDir = (Get-Location).Path
    }
}

$embedUrl = "https://www.python.org/ftp/python/$pythonVersion/python-$pythonVersion-embed-amd64.zip"
$zipFile = Join-Path $pythonDir "python-embed.zip"
$pythonExe = Join-Path $pythonDir "python.exe"
$pthFile = Join-Path $pythonDir "python311._pth"
$pipExe = Join-Path $pythonDir "Scripts\pip.exe"

Write-Host "=== Setting up embedded Python runtime ===" -ForegroundColor Cyan

if (-not (Test-Path $pythonExe)) {
    Write-Host "Downloading Python $pythonVersion Embedded Distribution..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $embedUrl -OutFile $zipFile -UseBasicParsing

    Write-Host "Extracting to $pythonDir..." -ForegroundColor Yellow
    Expand-Archive -Path $zipFile -DestinationPath $pythonDir -Force
    Remove-Item $zipFile -Force
} else {
    Write-Host "python.exe already exists, skipping download" -ForegroundColor Green
}

if (Test-Path $pthFile) {
    $pthContent = @(Get-Content $pthFile)
    if ($pthContent.Count -gt 0) {
        $pthContent[0] = $pthContent[0].TrimStart([char]0xFEFF)
    }
    $pthContent = $pthContent -replace "^#import site", "import site"
    if (-not ($pthContent -match "Lib\\site-packages")) {
        $pthContent += "Lib\site-packages"
    }
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($pthFile, [string[]]$pthContent, $utf8NoBom)
    Write-Host "Configured python311._pth for pip and site-packages" -ForegroundColor Green
}

if (-not (Test-Path $pipExe)) {
    Write-Host "Installing pip..." -ForegroundColor Yellow
    $getPipUrl = "https://bootstrap.pypa.io/get-pip.py"
    $getPipFile = Join-Path $pythonDir "get-pip.py"
    Invoke-WebRequest -Uri $getPipUrl -OutFile $getPipFile -UseBasicParsing

    & $pythonExe $getPipFile
    Remove-Item $getPipFile -Force -ErrorAction SilentlyContinue
    Write-Host "pip installed" -ForegroundColor Green
} else {
    Write-Host "pip already exists, skipping install" -ForegroundColor Green
}

Write-Host "Installing xlwings..." -ForegroundColor Yellow
& $pipExe install xlwings --no-warn-script-location
Write-Host "xlwings installed" -ForegroundColor Green

Write-Host "Verifying runtime..." -ForegroundColor Yellow
$result = & $pythonExe -c "import xlwings; print('xlwings version:', xlwings.__version__)"
if ($result -match "xlwings version") {
    Write-Host "Runtime verification succeeded: $result" -ForegroundColor Green
} else {
    Write-Host "Runtime verification failed" -ForegroundColor Red
    exit 1
}

Write-Host "Cleaning unnecessary files..." -ForegroundColor Yellow
$cleanupPatterns = @("*.chm", "*.h", "*.lib", "*.pdb", "*.pyd.orig")
foreach ($pattern in $cleanupPatterns) {
    Get-ChildItem -Path $pythonDir -Filter $pattern -Recurse |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

Write-Host "=== Embedded Python runtime setup complete ===" -ForegroundColor Cyan
$size = (Get-ChildItem -Path $pythonDir -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB
Write-Host "Total size: $([math]::Round($size, 1)) MB" -ForegroundColor Cyan
