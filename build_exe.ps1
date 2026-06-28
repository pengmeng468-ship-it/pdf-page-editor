$ErrorActionPreference = "Stop"
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = "C:\Users\23291\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$Venv = Join-Path $Here ".venv"

Set-Location $Here

if (!(Test-Path $Venv)) {
  & $Python -m venv $Venv
}

$VenvPython = Join-Path $Venv "Scripts\python.exe"
& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r requirements.txt

& $VenvPython -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --windowed `
  --name "PDFTool" `
  --add-data "static;static" `
  --hidden-import tkinter `
  app.py

& $VenvPython copy_release.py
Write-Host "DONE"
