# install.ps1 — One-liner installer for oa.exe
# Usage: powershell -ExecutionPolicy ByPass -c "irm https://raw.githubusercontent.com/albertxli/office-automation/main/install.ps1 | iex"

$installDir = "$env:LOCALAPPDATA\oa"
$exePath = "$installDir\oa.exe"
$url = "https://github.com/albertxli/office-automation/releases/latest/download/oa.exe"

Write-Host "Installing oa..." -ForegroundColor Cyan

# Create install directory
New-Item -Force -ItemType Directory $installDir | Out-Null

# Download binary
Write-Host "  Downloading oa.exe..."
Invoke-RestMethod $url -OutFile $exePath

# Add to PATH if not already there
$userPath = [Environment]::GetEnvironmentVariable("Path", "User")
if ($userPath -notlike "*$installDir*") {
    [Environment]::SetEnvironmentVariable("Path", "$userPath;$installDir", "User")
    $env:Path = "$env:Path;$installDir"
    Write-Host "  Added $installDir to PATH"
}

# Verify
$version = & $exePath --version 2>&1
Write-Host ""
Write-Host "  Installed: $version" -ForegroundColor Green
Write-Host "  Location:  $exePath" -ForegroundColor Green
Write-Host ""
Write-Host "  Restart your terminal, then run: oa --help" -ForegroundColor Yellow
