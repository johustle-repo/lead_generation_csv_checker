$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$appName = "ElmarsLeadGenerationQualityStudio"
$entryPoint = "check_nulls.py"
$distFolder = Join-Path "dist" $appName
$venvPath = Join-Path $projectRoot ".venv"
$pythonExe = Join-Path $venvPath "Scripts\python.exe"
$systemPython = "python"

function Invoke-Step {
    param(
        [string]$CommandLabel,
        [scriptblock]$Command
    )

    Write-Host $CommandLabel
    & $Command
    if ($LASTEXITCODE -ne 0) {
        throw "Step failed: $CommandLabel"
    }
}

if (-not (Test-Path $venvPath)) {
    Write-Host "Creating virtual environment..."
    & $systemPython -m venv $venvPath
}

if (-not (Test-Path $pythonExe)) {
    throw "Virtual environment Python was not created successfully."
}

$useSystemPython = $false
$venvProbeSucceeded = $true
try {
    & $pythonExe -c "import pip" 2>$null | Out-Null
    if ($LASTEXITCODE -ne 0) {
        $venvProbeSucceeded = $false
    }
}
catch {
    $venvProbeSucceeded = $false
}

if (-not $venvProbeSucceeded) {
    Write-Host "Local virtual environment pip is unavailable. Falling back to the system Python environment..."
    $useSystemPython = $true
}

$buildPython = if ($useSystemPython) { $systemPython } else { $pythonExe }

Invoke-Step "Installing build dependencies..." { & $buildPython -m pip install --upgrade pip }
Invoke-Step "Installing project requirements..." { & $buildPython -m pip install -r requirements.txt }

Invoke-Step "Building Windows executable..." {
    & $buildPython -m PyInstaller `
        --noconfirm `
        --clean `
        --windowed `
        --onedir `
        --contents-directory "." `
        --noupx `
        --icon "app-logo.ico" `
        --version-file "version_info.txt" `
        --add-data "app-logo.png;." `
        --name $appName `
        $entryPoint
}

Write-Host ""
Write-Host "Build complete."
Write-Host "Executable: $distFolder\$appName.exe"
