<#
.SYNOPSIS
    Build Memory Palace into a standalone Windows executable with PyInstaller.

.EXAMPLE
    .\build.ps1            # dist\MemoryPalace.exe  (single file)
    .\build.ps1 -OneDir    # dist\MemoryPalace\MemoryPalace.exe  (folder; starts faster)

    The exe keeps its library in a "data" folder next to itself, so put it
    somewhere writable (e.g. Documents\MemoryPalace), not in Program Files.
#>
param([switch]$OneDir)

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
    "--exclude-module", "fitz",         # legacy PyMuPDF alias; the app imports 'pymupdf'
    "main.py"
)

Write-Host "Building ($mode)... this takes a minute or two." -ForegroundColor Cyan
& $py -m PyInstaller @pyiArgs
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed" }

$exe = if ($OneDir) { "dist\MemoryPalace\MemoryPalace.exe" } else { "dist\MemoryPalace.exe" }
$sizeMb = [math]::Round((Get-Item $exe).Length / 1MB, 1)
Write-Host ""
Write-Host "Done: $exe ($sizeMb MB)" -ForegroundColor Green
Write-Host "Copy it anywhere writable and double-click it. Your library is stored in a 'data' folder beside it."
