# 🚀 Quick Start Script for Webhook Automation
# This script helps you get started with automated email-to-ticket system

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "🎫 EMAIL-TO-TICKET WEBHOOK AUTOMATION" -ForegroundColor Cyan
Write-Host "   Quick Start Script" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Check if in correct directory
if (-not (Test-Path ".\backend")) {
    Write-Host "❌ Error: Please run this script from the project root directory" -ForegroundColor Red
    Write-Host "   Current location: $(Get-Location)" -ForegroundColor Yellow
    Write-Host "   Expected: C:\Users\LAKSHYA\Desktop\email-summarizer" -ForegroundColor Yellow
    exit 1
}

Write-Host "📋 STEP 1: Checking Prerequisites" -ForegroundColor Green
Write-Host ""

# Check Python
try {
    $pythonVersion = python --version 2>&1
    Write-Host "✅ Python: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "❌ Python not found! Please install Python 3.8+" -ForegroundColor Red
    exit 1
}

# Check if ngrok is installed
try {
    $ngrokVersion = ngrok version 2>&1
    Write-Host "✅ ngrok: Installed" -ForegroundColor Green
} catch {
    Write-Host "⚠️  ngrok not found" -ForegroundColor Yellow
    Write-Host "   Install with: choco install ngrok" -ForegroundColor Yellow
    Write-Host "   Or download from: https://ngrok.com/download" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   Do you want to continue anyway? (y/n): " -NoNewline -ForegroundColor Yellow
    $continue = Read-Host
    if ($continue -ne 'y') {
        exit 1
    }
}

Write-Host ""
Write-Host "📦 STEP 2: Installing Python Dependencies" -ForegroundColor Green
Write-Host ""

Set-Location backend
pip install -r requirements.txt | Out-Null
if ($LASTEXITCODE -eq 0) {
    Write-Host "✅ Dependencies installed" -ForegroundColor Green
} else {
    Write-Host "❌ Failed to install dependencies" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "🔐 STEP 3: Configuration Check" -ForegroundColor Green
Write-Host ""

# Check .env file
if (Test-Path ".env") {
    $envContent = Get-Content ".env" -Raw
    
    $required = @{
        "CLIENT_ID" = "Azure App Client ID"
        "CLIENT_SECRET" = "Azure App Client Secret"
        "WEBHOOK_URL" = "Public webhook URL (ngrok)"
        "GEMINI_API_KEY" = "Google Gemini API Key"
    }
    
    $missing = @()
    
    foreach ($key in $required.Keys) {
        if ($envContent -match "$key=(.+)") {
            $value = $matches[1].Trim()
            if ($value) {
                Write-Host "✅ $key configured" -ForegroundColor Green
            } else {
                Write-Host "❌ $key not set" -ForegroundColor Red
                $missing += "$key ($($required[$key]))"
            }
        } else {
            Write-Host "❌ $key missing" -ForegroundColor Red
            $missing += "$key ($($required[$key]))"
        }
    }
    
    if ($missing.Count -gt 0) {
        Write-Host ""
        Write-Host "⚠️  Missing required configuration:" -ForegroundColor Yellow
        foreach ($item in $missing) {
            Write-Host "   - $item" -ForegroundColor Yellow
        }
        Write-Host ""
        Write-Host "📖 Please read: ..\WEBHOOK_AUTOMATION_GUIDE.md" -ForegroundColor Cyan
        Write-Host "   For complete setup instructions" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "   Do you want to continue anyway? (y/n): " -NoNewline -ForegroundColor Yellow
        $continue = Read-Host
        if ($continue -ne 'y') {
            exit 1
        }
    }
} else {
    Write-Host "❌ .env file not found!" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "🚀 READY TO START!" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "What would you like to do?" -ForegroundColor Yellow
Write-Host ""
Write-Host "1. Start ngrok tunnel (required for webhooks)" -ForegroundColor White
Write-Host "2. Start FastAPI backend server" -ForegroundColor White
Write-Host "3. Run webhook setup script" -ForegroundColor White
Write-Host "4. Start subscription renewal service" -ForegroundColor White
Write-Host "5. View complete guide" -ForegroundColor White
Write-Host "6. Exit" -ForegroundColor White
Write-Host ""
Write-Host "Enter choice (1-6): " -NoNewline -ForegroundColor Yellow
$choice = Read-Host

switch ($choice) {
    "1" {
        Write-Host ""
        Write-Host "🌐 Starting ngrok tunnel..." -ForegroundColor Green
        Write-Host ""
        Write-Host "Copy the HTTPS URL and add it to .env as WEBHOOK_URL" -ForegroundColor Yellow
        Write-Host "Example: WEBHOOK_URL=https://abc123.ngrok.io" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Press Ctrl+C to stop ngrok" -ForegroundColor Yellow
        Write-Host ""
        Start-Sleep -Seconds 2
        ngrok http 8001
    }
    "2" {
        Write-Host ""
        Write-Host "🚀 Starting FastAPI backend server..." -ForegroundColor Green
        Write-Host ""
        Write-Host "Server will run at: http://localhost:8001" -ForegroundColor Cyan
        Write-Host "API docs at: http://localhost:8001/docs" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Press Ctrl+C to stop server" -ForegroundColor Yellow
        Write-Host ""
        Start-Sleep -Seconds 2
        python main.py
    }
    "3" {
        Write-Host ""
        Write-Host "📡 Running webhook setup..." -ForegroundColor Green
        Write-Host ""
        Write-Host "⚠️  Make sure:" -ForegroundColor Yellow
        Write-Host "   1. ngrok is running" -ForegroundColor Yellow
        Write-Host "   2. FastAPI server is running" -ForegroundColor Yellow
        Write-Host "   3. WEBHOOK_URL is set in .env" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Continue? (y/n): " -NoNewline -ForegroundColor Yellow
        $confirm = Read-Host
        if ($confirm -eq 'y') {
            python setup_webhooks.py
        }
    }
    "4" {
        Write-Host ""
        Write-Host "🔄 Starting subscription renewal service..." -ForegroundColor Green
        Write-Host ""
        Write-Host "This will auto-renew your webhook subscription every 3 days" -ForegroundColor Cyan
        Write-Host "Press Ctrl+C to stop" -ForegroundColor Yellow
        Write-Host ""
        Start-Sleep -Seconds 2
        python subscription_renewal_service.py
    }
    "5" {
        Write-Host ""
        Write-Host "📖 Opening complete guide..." -ForegroundColor Green
        Set-Location ..
        if (Test-Path ".\WEBHOOK_AUTOMATION_GUIDE.md") {
            Start-Process ".\WEBHOOK_AUTOMATION_GUIDE.md"
            Write-Host "✅ Guide opened in default markdown viewer" -ForegroundColor Green
        } else {
            Write-Host "❌ Guide not found" -ForegroundColor Red
        }
    }
    "6" {
        Write-Host ""
        Write-Host "👋 Goodbye!" -ForegroundColor Cyan
        exit 0
    }
    default {
        Write-Host ""
        Write-Host "❌ Invalid choice" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
