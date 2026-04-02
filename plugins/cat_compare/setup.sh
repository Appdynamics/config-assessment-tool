# setup.sh

#!/bin/bash
set -e

# Detect a Python executable
PYTHON_BIN=""
if command -v python3 >/dev/null 2>&1; then
  PYTHON_BIN="python3"
elif command -v python >/dev/null 2>&1; then
  PYTHON_BIN="python"
else
  echo "Python is not installed. Please install Python 3.8+ and rerun."
  exit 1
fi

echo "Using Python at: $(command -v $PYTHON_BIN)"

# Create a virtual environment
echo "Creating a virtual environment..."
$PYTHON_BIN -m venv venv

# Activate the virtual environment (POSIX shells)
if [ -f "venv/bin/activate" ]; then
  echo "Activating the virtual environment..."
  # shellcheck disable=SC1091
  source venv/bin/activate
elif [ -f "venv/Scripts/activate" ]; then
  echo "Activating the virtual environment (Windows)..."
  # shellcheck disable=SC1091
  source venv/Scripts/activate
else
  echo "Could not find venv activation script."
  exit 1
fi

# Upgrade pip/setuptools/wheel and install requirements
echo "Upgrading pip, setuptools, and wheel..."
python -m pip install --upgrade pip setuptools wheel

echo "Installing required Python packages..."
python -m pip install -r requirements.txt

echo
echo "Setup complete."
echo "To run the app:"
echo "  Activate venv (macOS/Linux): source venv/bin/activate"
echo "  Activate venv (Windows, CMD/PowerShell): venv\\Scripts\\activate"
echo "  Start the app: python core.py"