$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

& "$projectRoot\build_exe.ps1"

$possibleCompilers = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
    "${env:LOCALAPPDATA}\Programs\Inno Setup 6\ISCC.exe"
)

$iscc = $possibleCompilers | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $iscc) {
    Write-Host ""
    Write-Host "Executable build completed, but Inno Setup was not found."
    Write-Host "Install Inno Setup 6, then rerun build_installer.ps1 to create a setup installer."
    exit 0
}

Write-Host "Building installer..."
& $iscc "installer.iss"

Write-Host ""
Write-Host "Installer complete."
Write-Host "Installer: installer-dist\ElmarLeadGenerationQualityStudioSetup.exe"
