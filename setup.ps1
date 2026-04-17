$VENV_DIR = ".venv"

Write-Host "[1/4] Creating virtual environment..."
python.exe -m venv $VENV_DIR

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