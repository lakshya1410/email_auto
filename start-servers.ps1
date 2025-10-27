# Email Summarizer - Start Both Servers# Start both frontend and backend servers

# This script starts the frontend and backend servers in separate windowsWrite-Host "Starting Email Summarizer servers..." -ForegroundColor Green



Write-Host "🚀 Starting Email Summarizer Servers..." -ForegroundColor Cyan# Start backend

Write-Host ""Start-Process pwsh -ArgumentList "-NoExit", "-Command", "cd 'C:\Users\LAKSHYA\Desktop\email-summarizer\backend'; uvicorn main:app --host 0.0.0.0 --port 8001 --ssl-keyfile certs/key.pem --ssl-certfile certs/cert.pem --reload"



# Check if ports are already in use# Wait a bit

$frontendPort = Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinueStart-Sleep -Seconds 2

$backendPort = Get-NetTCPConnection -LocalPort 8001 -State Listen -ErrorAction SilentlyContinue

# Start frontend

if ($frontendPort) {Start-Process pwsh -ArgumentList "-NoExit", "-Command", "cd 'C:\Users\LAKSHYA\Desktop\email-summarizer\frontend'; npm start"

    Write-Host "⚠️  Port 3000 is already in use!" -ForegroundColor Yellow

    Write-Host "   Frontend may already be running." -ForegroundColor YellowWrite-Host "Servers starting... Please wait for webpack to compile." -ForegroundColor Yellow

    Write-Host ""Write-Host "Backend: https://localhost:8001" -ForegroundColor Cyan

}Write-Host "Frontend: https://localhost:3000" -ForegroundColor Cyan


if ($backendPort) {
    Write-Host "⚠️  Port 8001 is already in use!" -ForegroundColor Yellow
    Write-Host "   Backend may already be running." -ForegroundColor Yellow
    Write-Host ""
}

# Start Frontend (Webpack Dev Server)
Write-Host "📦 Starting Frontend Server..." -ForegroundColor Green
Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$PSScriptRoot\frontend'; Write-Host '🌐 Frontend Server Starting...' -ForegroundColor Cyan; npm run dev"

# Wait a moment
Start-Sleep -Seconds 2

# Start Backend (FastAPI with Uvicorn)
Write-Host "🐍 Starting Backend Server..." -ForegroundColor Green
Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$PSScriptRoot\backend'; Write-Host '⚡ Backend Server Starting...' -ForegroundColor Cyan; python -m uvicorn main:app --host 0.0.0.0 --port 8001 --reload"

Write-Host ""
Write-Host "✅ Both servers are starting in separate windows!" -ForegroundColor Green
Write-Host ""
Write-Host "📍 Frontend: https://localhost:3000" -ForegroundColor Cyan
Write-Host "📍 Backend:  http://localhost:8001" -ForegroundColor Cyan
Write-Host ""
Write-Host "📖 Next steps:" -ForegroundColor Yellow
Write-Host "   1. Wait for both servers to fully start" -ForegroundColor White
Write-Host "   2. Open Outlook Desktop" -ForegroundColor White
Write-Host "   3. Install add-in from: frontend\manifest.xml" -ForegroundColor White
Write-Host "   4. Select an email and click 'Summarize Email'" -ForegroundColor White
Write-Host ""
Write-Host "Press any key to exit this window..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
