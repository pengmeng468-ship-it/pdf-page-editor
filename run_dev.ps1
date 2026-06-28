$ErrorActionPreference = "Stop"
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = "C:\Users\23291\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
Set-Location $Here
& $Python app.py
