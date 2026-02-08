# Windows Production Deployment Script
# Run as Administrator in PowerShell
# Usage: .\deploy-windows.ps1

$ErrorActionPreference = "Stop"

Write-Host "🚀 Starting Windows production deployment..." -ForegroundColor Green

# Check if running as Administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "❌ Please run as Administrator" -ForegroundColor Red
    exit 1
}

# 1. Install Python dependencies
Write-Host "📦 Installing Python dependencies..." -ForegroundColor Yellow
pip install --upgrade pip
pip install -r requirements.txt

# 2. Test the application
Write-Host "🧪 Testing application..." -ForegroundColor Yellow
python -c "from api import app; print('✅ Application imports successfully')"

# 3. Create Windows service using NSSM (optional)
Write-Host ""
Write-Host "📝 To run as a Windows Service:" -ForegroundColor Cyan
Write-Host "1. Install NSSM: choco install nssm (requires Chocolatey)"
Write-Host "   or download from https://nssm.cc/download"
Write-Host ""
Write-Host "2. Create service:"
Write-Host '   nssm install SHMReports "C:\Python39\python.exe"'
Write-Host '   nssm set SHMReports AppDirectory "' + (Get-Location).Path + '"'
Write-Host '   nssm set SHMReports AppParameters "run.py serve --production --port 5000"'
Write-Host '   nssm start SHMReports'
Write-Host ""

# 4. For IIS deployment
Write-Host "🌐 For IIS deployment with SSL:" -ForegroundColor Cyan
Write-Host "1. Install IIS and URL Rewrite module"
Write-Host "2. Create a new site in IIS"
Write-Host "3. Configure SSL certificate in IIS"
Write-Host "4. Add reverse proxy rule to forward to http://localhost:5000"
Write-Host ""

# 5. Direct SSL deployment
Write-Host "🔒 For direct SSL deployment:" -ForegroundColor Cyan
Write-Host '1. Place your SSL certificate files in a secure location'
Write-Host '2. Run: python run.py serve --production --port 443 --ssl-cert "C:\Path\To\cert.crt" --ssl-key "C:\Path\To\key.key"'
Write-Host ""

# 6. Start production server (HTTP - use IIS for SSL)
Write-Host "▶️  Starting production server..." -ForegroundColor Yellow
Write-Host "Starting on http://0.0.0.0:5000"
Write-Host "Press Ctrl+C to stop"
Write-Host ""

python run.py serve --production --port 5000
