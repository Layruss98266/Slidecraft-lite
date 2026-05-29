# SlideCraft Lite launcher (PowerShell)
$ErrorActionPreference = "Stop"
Set-Location -LiteralPath (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not (Test-Path ".venv")) {
    Write-Host ">> First run: creating .venv"
    python -m venv .venv
}
. .\.venv\Scripts\Activate.ps1
$flaskInstalled = $false
try { python -c "import flask" 2>$null; if ($LASTEXITCODE -eq 0) { $flaskInstalled = $true } } catch {}
if (-not $flaskInstalled) {
    Write-Host ">> Installing dependencies (one-time, ~30 s)"
    pip install --upgrade pip
    pip install -r requirements.txt
}
$soffice = Get-Command soffice -ErrorAction SilentlyContinue
if (-not $soffice -and -not (Test-Path "C:\Program Files\LibreOffice\program\soffice.exe")) {
    Write-Warning "LibreOffice not found — PPTX conversion will use a low-fidelity fallback."
    Write-Warning "Install from https://libreoffice.org/download"
}
python app.py
