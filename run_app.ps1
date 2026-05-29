# All In One Validation Formatter - Easy Launcher
Write-Host "Starting All In One Validation Formatter..." -ForegroundColor Cyan
Set-Location $PSScriptRoot

# Check if streamlit is installed
try {
    python -c "import streamlit" 2>$null
} catch {
    Write-Host "Streamlit not found. Installing dependencies..." -ForegroundColor Yellow
    pip install -r requirements.txt
}

python -m streamlit run app.py
