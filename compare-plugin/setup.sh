#!/bin/bash

# Function to check if a command exists
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# Check if Python 3 is installed
if command_exists python3; then
    echo "Python 3 is installed."
else
    echo "Python 3 is not installed. Please install Python 3 before proceeding."
    exit 1
fi

# Create a virtual environment
echo "Creating a virtual environment..."
python3 -m venv venv

# Activate the virtual environment
echo "Activating the virtual environment..."
source venv/bin/activate

# Install the required packages
echo "Installing required Python packages..."
pip install -r requirements.txt

echo "Setup is complete. You can now run the script using:"
echo "  source venv/bin/activate"
echo "  python3 compare_results/core.py"

