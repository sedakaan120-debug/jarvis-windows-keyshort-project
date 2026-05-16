param(
    [switch]$Run
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

Write-Host "J.A.R.V.I.S Windows setup" -ForegroundColor Cyan

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    throw "Python bulunamadi. Python 3.11+ kurup PATH'e ekleyin."
}

if (-not (Test-Path ".venv")) {
    python -m venv .venv
}

& ".\.venv\Scripts\python.exe" -m pip install --upgrade pip
& ".\.venv\Scripts\python.exe" -m pip install -r requirements.txt

Write-Host "Kurulum tamamlandi." -ForegroundColor Green
Write-Host "Baslatmak icin: .\.venv\Scripts\python.exe main.py"

if ($Run) {
    & ".\.venv\Scripts\python.exe" main.py
}
