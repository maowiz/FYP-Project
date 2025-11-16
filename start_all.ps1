param(
    [switch]$NoFrontendDelay
)

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "Starting backend (python server.py)..."
Start-Process powershell -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-Command","cd `"$repoRoot`"; python server.py" | Out-Null

if (-not $NoFrontendDelay) {
    Start-Sleep -Seconds 2
}

Write-Host "Starting frontend dev server..."
Push-Location (Join-Path $repoRoot 'frontend advance')
npm run dev
Pop-Location
