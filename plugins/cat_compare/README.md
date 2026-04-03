# CompareResults

**CompareResults** is a plugin for the [Configuration Assessment Tool (CAT)](https://github.com/Appdynamics/config-assessment-tool). 

It compares **Previous** and **Current** CAT workbooks for **APM**, **BRUM**, and **MRUM**, providing insights into how your application instrumentation has evolved.

## Key Features
CompareResults generates three outputs:
- **Excel Comparison Workbook**: A detailed, low-level comparison of the two workbooks.
- **PowerPoint Deck**: A high-level summary of the comparison results.
- **JSON Snapshot**: Used by the **Insights** view in the UI for further analysis.

--------------------------------------------------------------------------------

## Supported Files

You can compare CAT workbooks with the following naming conventions:
- `*-MaturityAssessment-apm.xlsx`
- `*-MaturityAssessment-brum.xlsx`
- `*-MaturityAssessment-mrum.xlsx`

### Important Rules
- Both the **Previous** and **Current** workbooks must be from the **same Controller**.
- The **Previous** workbook should be dated earlier than the **Current** workbook (recommended).
- Ensure you select the correct domain (APM, BRUM, or MRUM) when uploading files.

--------------------------------------------------------------------------------

## Requirements

- Python 3.x
- Dependencies listed in `requirements.txt`
- Microsoft Excel installed (required for formula recalculation via `xlwings`)
  - macOS: Excel for Mac installed and can open normally
  - Windows: Desktop Excel installed and can open normally

Optional / Notes:
- Internet access is only needed if your UI loads external assets (e.g., Chart.js via CDN). If you bundle assets locally, internet is not required.

--------------------------------------------------------------------------------

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/Appdynamics/config-assessment-tool.git
   cd config-assessment-tool/compare-plugin
   ```

--------------------------------------------------------------------------------

## Quick Start (Recommended) — One-command launcher

The launcher creates a virtual environment (`.venv`), installs requirements, starts the web UI, and opens your browser.

macOS / Linux:
    cd path/to/config-assessment-tool/compare-plugin
    python3 run_tool.py

Windows (PowerShell):
    cd path\to\config-assessment-tool\compare-plugin
    python .\run_tool.py

Windows (Command Prompt):
    cd path\to\config-assessment-tool\compare-plugin
    python run_tool.py

When it starts, you should see something like:
- Starting Config Assessment Tool on http://127.0.0.1:5000 ...

The UI should open automatically. If it doesn’t, open:
- http://127.0.0.1:5000/

Stop the server with:
- macOS/Linux: Ctrl + C
- Windows: Ctrl + C

--------------------------------------------------------------------------------

## One-click / Double-click launcher

If you want a “double-click feel”, you can run the launcher script directly (it will create `.venv`, install requirements, start the web UI, and open your browser).

### macOS / Linux (double-click or run from Terminal)

**Option A — Terminal (recommended the first time):**
1) Go to the `compare-plugin` folder:
    cd path/to/config-assessment-tool/compare-plugin

2) One-time: allow the launcher to run:
    chmod +x run_tool.py

3) Run it:
    ./run_tool.py
   OR:
    python3 run_tool.py

**Option B — Finder double-click (works if macOS allows it):**
- You can try double-clicking `run_tool.py`, but macOS may open it in an editor instead of running it.
- If that happens, use Option A.

Notes:
- If macOS blocks execution (Gatekeeper), right-click the file → **Open** (once), or run via Terminal.
- You may be prompted for permissions the first time Excel automation runs.

### Windows (double-click)

Windows will usually open `.py` files with Python if Python is installed and associated with `.py`.

**Option A — Double-click**
- Double-click `run_tool.py` inside `compare-plugin`.

**Option B — Right-click**
- Right-click `run_tool.py` → **Open with** → **Python** (or “Python Launcher”).

If double-click does nothing, use PowerShell:
    cd path\to\config-assessment-tool\compare-plugin
    python .\run_tool.py

--------------------------------------------------------------------------------


## Using the UI

1) Open the UI:
- http://127.0.0.1:5000/

2) Select the domain:
- APM, BRUM, or MRUM

3) Upload:
- Previous CAT workbook (older)
- Current CAT workbook (newer)

4) Click Upload and Compare (or the equivalent button)

You’ll then see links to download:
- Excel comparison workbook
- PowerPoint deck
- JSON snapshot (used by Insights)

--------------------------------------------------------------------------------

## Outputs

Outputs are written to the results/output folder configured by the app (commonly `compare-plugin/results/`).

Typical filenames:
- `Analysis_Summary_APM.xlsx` and `Analysis_Summary_APM.pptx`
- `Analysis_Summary_BRUM.xlsx` and `Analysis_Summary_BRUM.pptx`
- `Analysis_Summary_MRUM.xlsx` and `Analysis_Summary_MRUM.pptx`
- `analysis_summary_<domain>_<timestamp>.json`

What each output is for:
- Excel: deep-dive comparison (per sheet / per metric)
- PowerPoint: leadership summary + key callouts + deep-dive slides
- JSON: powers the UI “Insights” view (per-app drilldown)

--------------------------------------------------------------------------------

## Troubleshooting

1) “UI opens but Insights is blank / missing comparisons”
- Ensure the comparison workbook has an Analysis sheet with a `name` column.
- Ensure domain columns include strings like “Upgraded” / “Downgraded” (Insights detects these).
- Ensure the JSON is being generated and saved into the configured results folder.

2) ModuleNotFoundError: No module named 'compare_tool'
This usually means you ran a file directly instead of running from the plugin root.
- Always run from the `compare-plugin` folder using:
    python3 run_tool.py
  Or on Windows:
    python run_tool.py

If you are intentionally running the Flask app as a module, do it from the plugin root:
    python -m webapp.app

3) can't open file ... app.py: [Errno 2]
Your `app.py` lives under `webapp/`. Don’t run `python app.py` from the repo root.
Use:
    python -m webapp.app
or (recommended):
    python run_tool.py

4) Import "compare_tool.logging_config" could not be resolved
This is an editor/Pylance warning if you removed/renamed the file but still import it.
Fix by either:
- removing the import where it’s referenced, OR
- restoring the file if you still want centralized logging setup

5) threading is not defined
Add the missing import at the top of the file that uses it:
    import threading

6) Excel / xlwings issues
Because the tool may open Excel to recalc formulas:
- Make sure Excel is installed and opens normally
- Close any “blocked” Excel dialogs (first-run prompts, file recovery, etc.)
- On macOS, Excel automation may require permissions the first time

7) Controller mismatch
If Previous + Current are from different Controllers, the tool will stop.
Re-run with two workbooks from the same Controller.

--------------------------------------------------------------------------------

## Support / Notes

- This tool compares CAT outputs and generates a “what changed” story between two points in time.
- For best results:
  - Keep CAT outputs named consistently
  - Run CAT with the same settings on both runs
  - Use a meaningful time gap between Previous and Current
- If you hit an issue, capture:
  - terminal logs
  - the two input workbook filenames
  - the generated comparison workbook (if it exists)
  and open a GitHub issue / share with the maintainer.
