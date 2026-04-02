#!/usr/bin/env python
"""
One-click launcher for the Config Assessment Tool.

- Creates a .venv if it doesn't exist
- Installs requirements
- Starts the Flask UI (webapp.app)
"""

import os
import sys
import subprocess
from pathlib import Path


ROOT = Path(__file__).resolve().parent             # compare-plugin/
VENV_DIR = ROOT / ".venv"
REQUIREMENTS = ROOT / "requirements.txt"


def get_venv_python() -> Path:
    if os.name == "nt":  # Windows
        return VENV_DIR / "Scripts" / "python.exe"
    else:                # macOS / Linux
        return VENV_DIR / "bin" / "python"


def ensure_venv():
    venv_python = get_venv_python()
    if not VENV_DIR.exists() or not venv_python.exists():
        if VENV_DIR.exists():
            print("Removing incomplete virtual environment...")
            import shutil
            shutil.rmtree(VENV_DIR)
        print("Creating virtual environment in .venv ...")
        subprocess.check_call([sys.executable, "-m", "venv", str(VENV_DIR)])
    else:
        print("Using existing virtual environment .venv")


def ensure_requirements(venv_python: Path):
    if REQUIREMENTS.exists():
        print("Installing dependencies from requirements.txt ...")
        subprocess.check_call([str(venv_python), "-m", "pip", "install", "--upgrade", "pip"])
        subprocess.check_call([str(venv_python), "-m", "pip", "install", "-r", str(REQUIREMENTS)])
    else:
        print("WARNING: requirements.txt not found; skipping dependency install.")


def main():
    os.chdir(ROOT)

    ensure_venv()
    venv_python = get_venv_python()
    if not venv_python.exists():
        raise SystemExit(f"Could not find venv Python at: {venv_python}")

    ensure_requirements(venv_python)

    # Run the Flask app as a module so `compare_tool` imports work
    print("Starting Config Assessment Tool on http://127.0.0.1:5000 ...")
    subprocess.check_call([str(venv_python), "-m", "webapp.app"])


if __name__ == "__main__":
    main()
