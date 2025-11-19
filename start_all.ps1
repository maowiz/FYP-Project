$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "Launching FYP System..." -ForegroundColor Cyan

# 1. Start Backend
Write-Host "Starting Backend Server..."
Start-Process powershell -ArgumentList "-NoExit", "-Command", "& {
    `$Host.UI.RawUI.WindowTitle = 'FYP BACKEND';
    cd '$repoRoot';
    Write-Host 'Starting Python Server...' -ForegroundColor Green;
    python server.py;
}"

# 2. Start Frontend
Write-Host "Starting Frontend Dev Server..."
Start-Process powershell -ArgumentList "-NoExit", "-Command", "& {
    `$Host.UI.RawUI.WindowTitle = 'FYP FRONTEND';
    cd '$repoRoot\frontend advance';
    Write-Host 'Starting React Dev Server...' -ForegroundColor Green;
    npm run dev;
}"

Write-Host "All systems launched!" -ForegroundColor Green

