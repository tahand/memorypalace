<#
.SYNOPSIS
    Build Memory Palace into a standalone Windows executable with PyInstaller.

.EXAMPLE
    .\build.ps1            # dist\MemoryPalace.exe  (single file)
    .\build.ps1 -OneDir    # dist\MemoryPalace\MemoryPalace.exe  (folder; starts faster)

    .\build.ps1 -NoDeploy  # build only, don't copy it anywhere

    After building, the exe is copied to -DeployTo (default
    OneDrive\Documents\MemoryPalace in your user folder). The exe keeps its
    library in a "data" folder next to itself, so that folder must be
    writable (not Program Files); the copy never touches "data".
#>
param(
    [switch]$OneDir,
    [switch]$NoDeploy,
    [string]$DeployTo = (Join-Path $env:USERPROFILE "OneDrive\Documents\MemoryPalace")
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$py = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) {
    Write-Host "Creating virtual environment..." -ForegroundColor Cyan
    python -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw "Could not create .venv. Is Python on PATH?" }
}
Write-Host "Installing requirements..." -ForegroundColor Cyan
& $py -m pip install --quiet --disable-pip-version-check -r requirements.txt
if ($LASTEXITCODE -ne 0) { throw "pip install -r requirements.txt failed" }

Write-Host "Installing PyInstaller..." -ForegroundColor Cyan
& $py -m pip install --quiet --disable-pip-version-check pyinstaller
if ($LASTEXITCODE -ne 0) { throw "Could not install PyInstaller" }

& $py tools\make_icon.py build\icon.ico
if ($LASTEXITCODE -ne 0) { throw "Icon generation failed" }

$mode = if ($OneDir) { "--onedir" } else { "--onefile" }
$pyiArgs = @(
    "--noconfirm", "--clean",
    "--windowed",                       # no console window
    $mode,
    "--name", "MemoryPalace",
    "--icon", "build\icon.ico",
    "--collect-data", "customtkinter",  # themes/fonts that customtkinter loads from disk
    "--collect-submodules", "winrt",    # Windows OCR bindings (imported lazily)
    "--hidden-import", "win32com.client",  # SAPI read-aloud voices
    "--exclude-module", "pytest",
    "--collect-submodules", "pdf2docx", # PDF -> Word conversion (imports PyMuPDF as 'fitz')
    "main.py"
)

Write-Host "Building ($mode)... this takes a minute or two." -ForegroundColor Cyan
& $py -m PyInstaller @pyiArgs
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed" }

$exe = if ($OneDir) { "dist\MemoryPalace\MemoryPalace.exe" } else { "dist\MemoryPalace.exe" }
$sizeMb = [math]::Round((Get-Item $exe).Length / 1MB, 1)
Write-Host ""
Write-Host "Done: $exe ($sizeMb MB)" -ForegroundColor Green

if ($NoDeploy) {
    Write-Host "Copy it anywhere writable and double-click it. Your library is stored in a 'data' folder beside it."
    return
}
# Copy the new build to where the app is used. The 'data' folder there (your library) is left alone.
Write-Host "Copying to $DeployTo ..." -ForegroundColor Cyan
New-Item -ItemType Directory -Force $DeployTo | Out-Null
try {
    if ($OneDir) {
        Copy-Item "dist\MemoryPalace\*" $DeployTo -Recurse -Force
    } else {
        Copy-Item $exe $DeployTo -Force
    }
} catch {
    throw "Could not copy to $DeployTo. Is Memory Palace running from there? Close it and run the build again. ($_)"
}
Write-Host "Copied: $(Join-Path $DeployTo 'MemoryPalace.exe')" -ForegroundColor Green
