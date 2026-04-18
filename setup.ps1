$VENV_DIR = ".venv"

$python = Get-Command py -ErrorAction SilentlyContinue

# Fallback to python if py is not found
if (-not $python) {
    $python = Get-Command python -ErrorAction SilentlyContinue
}

# Fallback to python3 if python is not found
if (-not $python) {
    $python = Get-Command python3 -ErrorAction SilentlyContinue
}

if (-not $python) {
    Write-Error "Python not found. Please install Python v3 and ensure it's in PATH."
    exit 1
}

Write-Host "[1/4] Creating virtual environment..."
& "$python" -m venv $VENV_DIR

Write-Host "[2/4] Upgrading pip (inside venv)..."
& "$VENV_DIR\Scripts\python.exe" -m pip install --upgrade pip

Write-Host "[3/4] Installing dependencies..."
& "$VENV_DIR\Scripts\pip.exe" install -r requirements.txt

Write-Host "[4/4] Activating virtual environment..."
& "$VENV_DIR\Scripts\Activate.ps1"

Write-Host ""
Write-Host "Setup complete."
Write-Host ""
Write-Host "Run with: python -m xls2csv.cli --help"