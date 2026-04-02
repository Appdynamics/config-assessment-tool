# compare_tool/config.py

"""
config.py
---------
This module loads and parses the application configuration.

Purpose:
- Reads the `config.json` file to retrieve application settings.
- Provides configuration values such as `upload_folder` and `result_folder`.

Key Features:
- Ensures configuration values are accessible throughout the application.
"""

import json
from pathlib import Path
from typing import Optional, Dict, Any

# .../compare-plugin
BASE_DIR = Path(__file__).resolve().parent.parent


def load_config(config_path: Optional[str] = None) -> Dict[str, Any]:
    """
    Load config.json and normalise folder paths and template paths.
    """
    if config_path is None:
        cfg_path = BASE_DIR / "config.json"
    else:
        cfg_path = Path(config_path)

    if not cfg_path.exists():
        raise FileNotFoundError(f"Config file not found: {cfg_path}")

    with cfg_path.open() as f:
        cfg: Dict[str, Any] = json.load(f)

    # Normalise folders: upload, result, template
    for key in ("upload_folder", "result_folder", "TEMPLATE_FOLDER"):
        if key in cfg:
            cfg[key] = str((BASE_DIR / cfg[key]).resolve())

    template_folder = Path(cfg.get("TEMPLATE_FOLDER", BASE_DIR / "templates"))

    # Build per-domain template paths (APM/BRUM/MRUM)
    apm_tpl  = cfg.get("apm_template_file")
    brum_tpl = cfg.get("brum_template_file")
    mrum_tpl = cfg.get("mrum_template_file")

    cfg["apm_template_path"] = str(template_folder / apm_tpl) if apm_tpl else None
    cfg["brum_template_path"] = str(template_folder / brum_tpl) if brum_tpl else None
    cfg["mrum_template_path"] = str(template_folder / mrum_tpl) if mrum_tpl else None

    return cfg
