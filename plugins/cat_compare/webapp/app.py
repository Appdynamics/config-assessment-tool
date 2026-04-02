"""
app.py
------
This is the main entry point for the Flask web application.

Purpose:
- Initializes the Flask app.
- Sets up routes for the homepage, insights, and file uploads.
- Configures logging and application folders.

Key Routes:
- `/`: Renders the homepage.
- `/insights`: Renders the insights page.
- `/upload`: Handles file uploads for APM comparisons.
"""

import os
import json
import threading
import webbrowser
import datetime as dt
from pathlib import Path
from typing import List, Tuple
from flask import Flask, request, jsonify, render_template, send_from_directory
from typing import Optional
from compare_tool.config import load_config
from compare_tool.logging_config import setup_logging
from compare_tool.insights import build_comparison_json
from compare_tool.service import (
    run_comparison,        # APM
    run_comparison_brum,   # BRUM
    run_comparison_mrum,   # MRUM
    find_best_matching_files,  # Folder processing
    save_matched_files,        # Folder processing
)
import logging


def ts_now():
    return dt.datetime.now(dt.timezone.utc).strftime("%Y%m%d_%H%M%S")


setup_logging()
logging.info("Logging setup is complete.")

BASE_DIR = Path(__file__).resolve().parent.parent  # points at compare-plugin
config = load_config(str(BASE_DIR / "config.json"))

UPLOAD_FOLDER = config["upload_folder"]   # e.g. "uploads"
RESULT_FOLDER = config["result_folder"]   # e.g. "results"
HISTORY_FOLDER = BASE_DIR / "history"     # used by insights APIs

app = Flask(
    __name__,
    static_folder=str(BASE_DIR / "static"),
    template_folder=str(BASE_DIR / "templates"),
)

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)
os.makedirs(HISTORY_FOLDER, exist_ok=True)


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/insights", methods=["GET"])
def insights():
    return render_template("insights.html")


# ---------- APM upload (uses new compare_tool.service) -----------------------
@app.route("/upload", methods=["POST"])
def upload_apm():
    if "previous_file" not in request.files or "current_file" not in request.files:
        return render_template("index.html", message="Missing files."), 400

    prev = request.files["previous_file"]
    curr = request.files["current_file"]

    if not prev.filename or not curr.filename:
        return render_template("index.html", message="Please select both files."), 400

    prev_path = os.path.join(UPLOAD_FOLDER, "previous_apm.xlsx")
    curr_path = os.path.join(UPLOAD_FOLDER, "current_apm.xlsx")

    prev.save(prev_path)
    curr.save(curr_path)

    # Run the APM comparison pipeline
    output_file, ppt_file = run_comparison(
        previous_file_path=prev_path,
        current_file_path=curr_path,
        config=config,
    )

    # 🔹 Build JSON snapshot in RESULT_FOLDER so /api/history can see it
    json_path, json_name, _ = build_comparison_json(
        domain="APM",
        comparison_result_path=output_file,
        current_file_path=curr_path,
        previous_file_path=prev_path,
        result_folder=RESULT_FOLDER,   # IMPORTANT
    )

    app.config["LAST_RESULT_APM"] = output_file
    app.config["LAST_PPT_APM"] = ppt_file
    app.config["LAST_JSON_APM"] = json_path

    msg = (
        "APM comparison completed. "
        f"Download Excel <a href='/download/{os.path.basename(output_file)}' style='color:#32CD32;'>here</a> "
        f"and PowerPoint <a href='/download/{os.path.basename(ppt_file)}' style='color:#32CD32;'>here</a>. "
        "Insights snapshot has been generated and will be available on the Insights page."
    )
    return render_template("index.html", message=msg)




@app.route("/download/<filename>")
def download(filename):
    return send_from_directory(RESULT_FOLDER, filename, as_attachment=True)


# ---------- BRUM / MRUM upload placeholders ---------------------------------
@app.route("/upload_brum", methods=["POST"])
def upload_brum():
    if "previous_brum" not in request.files or "current_brum" not in request.files:
        return render_template("index.html", message="Missing BRUM files."), 400

    prev = request.files["previous_brum"]
    curr = request.files["current_brum"]

    if not prev.filename or not curr.filename:
        return render_template("index.html", message="Please select both BRUM files."), 400

    prev_path = os.path.join(UPLOAD_FOLDER, "previous_brum.xlsx")
    curr_path = os.path.join(UPLOAD_FOLDER, "current_brum.xlsx")

    prev.save(prev_path)
    curr.save(curr_path)

    output_file, ppt_file = run_comparison_brum(
        previous_file_path=prev_path,
        current_file_path=curr_path,
        config=config,
    )

    json_path, json_name, _ = build_comparison_json(
        domain="BRUM",
        comparison_result_path=output_file,
        current_file_path=curr_path,
        previous_file_path=prev_path,
        result_folder=RESULT_FOLDER,
    )

    msg = (
        "BRUM comparison completed. "
        f"Download Excel <a href='/download/{os.path.basename(output_file)}' style='color:#32CD32;'>here</a> "
        f"and PowerPoint <a href='/download/{os.path.basename(ppt_file)}' style='color:#32CD32;'>here</a>. "
        "BRUM Insights snapshot has been generated and will be available on the Insights page."
    )
    return render_template("index.html", message=msg)


@app.route("/upload_mrum", methods=["POST"])
def upload_mrum():
    if "previous_mrum" not in request.files or "current_mrum" not in request.files:
        return render_template("index.html", message="Missing MRUM files."), 400

    prev = request.files["previous_mrum"]
    curr = request.files["current_mrum"]

    if not prev.filename or not curr.filename:
        return render_template("index.html", message="Please select both MRUM files."), 400

    prev_path = os.path.join(UPLOAD_FOLDER, "previous_mrum.xlsx")
    curr_path = os.path.join(UPLOAD_FOLDER, "current_mrum.xlsx")

    prev.save(prev_path)
    curr.save(curr_path)

    output_file, ppt_file = run_comparison_mrum(
        previous_file_path=prev_path,
        current_file_path=curr_path,
        config=config,
    )

    json_path, json_name, _ = build_comparison_json(
        domain="MRUM",
        comparison_result_path=output_file,
        current_file_path=curr_path,
        previous_file_path=prev_path,
        result_folder=RESULT_FOLDER,
    )

    msg = (
        "MRUM comparison completed. "
        f"Download Excel <a href='/download/{os.path.basename(output_file)}' style='color:#32CD32;'>here</a> "
        f"and PowerPoint <a href='/download/{os.path.basename(ppt_file)}' style='color:#32CD32;'>here</a>. "
        "MRUM Insights snapshot has been generated and will be available on the Insights page."
    )
    return render_template("index.html", message=msg)


# ---------- Folder upload (processes multiple data types) --------------------
@app.route("/upload_folders", methods=["POST"])
def upload_folders():
    logging.debug("[FOLDERS] Request files: %s", list(request.files.keys()))
    
    # Check if folders were uploaded
    if 'previous_folder' not in request.files or 'current_folder' not in request.files:
        logging.error("[FOLDERS] No folder part")
        return render_template('index.html', message="Error: Please select both previous and current folders."), 400
    
    # Get the selected data types from checkboxes
    selected_types = request.form.getlist('data_types')
    if not selected_types:
        return render_template('index.html', message="Error: Please select at least one data type (APM, BRUM, or MRUM)."), 400
    
    logging.info(f"[FOLDERS] Selected data types: {selected_types}")
    
    # Get all files from both folders
    previous_files = request.files.getlist('previous_folder')
    current_files = request.files.getlist('current_folder')
    
    logging.info(f"[FOLDERS] Previous folder: {len(previous_files)} files")
    logging.info(f"[FOLDERS] Current folder: {len(current_files)} files")
    
    # Find matching files for each data type
    matches = find_best_matching_files(previous_files, current_files)
    
    # Process each selected data type
    results = {}
    errors = []
    
    for data_type in selected_types:
        domain = data_type.upper()
        logging.info(f"[FOLDERS] Processing {domain}")
        
        try:
            # Save matched files for this domain
            previous_path, current_path = save_matched_files(matches, UPLOAD_FOLDER, data_type)
            
            if not previous_path or not current_path:
                errors.append(f"No matching {domain} files found in the selected folders.")
                continue
            
            # Process based on data type
            if data_type == 'apm':
                output_file, ppt_file = run_comparison(
                    previous_file_path=previous_path,
                    current_file_path=current_path,
                    config=config,
                )
            elif data_type == 'brum':
                output_file, ppt_file = run_comparison_brum(
                    previous_file_path=previous_path,
                    current_file_path=current_path,
                    config=config,
                )
            elif data_type == 'mrum':
                output_file, ppt_file = run_comparison_mrum(
                    previous_file_path=previous_path,
                    current_file_path=current_path,
                    config=config,
                )
            
            # Build JSON snapshot for insights
            json_path, json_name, _ = build_comparison_json(
                domain=domain,
                comparison_result_path=output_file,
                current_file_path=current_path,
                previous_file_path=previous_path,
                result_folder=RESULT_FOLDER,
            )
            
            # Store results
            results[domain] = {
                'xlsx': os.path.basename(output_file),
                'pptx': os.path.basename(ppt_file),
                'json': json_name
            }
            
            logging.info(f"[FOLDERS] Successfully processed {domain}")
            
        except Exception as e:
            logging.error(f"[FOLDERS] Error processing {domain}: {e}", exc_info=True)
            errors.append(f"{domain}: Error during processing - {str(e)}")
    
    # Generate response message
    if results:
        message_parts = ["Processing completed successfully!<br><br>"]
        
        for domain, files in results.items():
            message_parts.append(f"<strong>{domain}:</strong><br>")
            message_parts.append(f"• Results: <a href='/download/{files['xlsx']}' style='color: #32CD32;'>Download Excel</a><br>")
            message_parts.append(f"• PowerPoint: <a href='/download/{files['pptx']}' style='color: #32CD32;'>Download PPT</a><br>")
            message_parts.append(f"• JSON: <a href='/download/{files['json']}' style='color: #32CD32;'>Download JSON</a><br><br>")
        
        if errors:
            message_parts.append("<br><strong>Warnings:</strong><br>")
            for error in errors:
                message_parts.append(f"• {error}<br>")
        
        message = "".join(message_parts)
    else:
        message = f"Error: No files could be processed. Issues encountered:<br>{'<br>'.join(errors)}"
        return render_template('index.html', message=message), 400
    
    return render_template('index.html', message=message)


#####################################################################################
############## Utility for Index on Read (compare multiple output) ##################
#####################################################################################


def _slug(s: Optional[str]) -> str:
    if not s:
        return ""
    return "".join(ch.lower() for ch in s if ch.isalnum())


def scan_runs(folder: str, domain: str, controller_filter: Optional[str], limit: int):
    """
    Scan RESULT_FOLDER for analysis_summary_<domain>_*.json
    and build a list of run dicts for trends.
    """
    prefix = f"analysis_summary_{domain}_"
    if not os.path.isdir(folder):
        return []

    runs = []
    for name in sorted(os.listdir(folder), reverse=True):
        if not (name.startswith(prefix) and name.endswith(".json")):
            continue

        path = os.path.join(folder, name)
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            meta = payload.get("meta") or {}
        except Exception:
            continue

        controller = meta.get("controller")
        if controller_filter and _slug(controller) != _slug(controller_filter):
            continue

        previousDate = meta.get("previousDate") or ""
        currentDate = meta.get("currentDate") or ""
        compareDate = meta.get("compareDate") or ""

        improved = int(meta.get("improved", 0))
        degraded = int(meta.get("degraded", 0))
        percentage = float(meta.get("percentage", 0.0))
        tiers = meta.get("tiers") or {}

        runs.append(
            {
                "file": name,
                "controller": controller,
                "previousDate": previousDate,
                "currentDate": currentDate,
                "compareDate": compareDate,
                "improved": improved,
                "degraded": degraded,
                "percentage": percentage,
                "tiers": tiers,
                "sortPrev": previousDate,
            }
        )

    # newest compareDate first
    runs.sort(key=lambda r: r["compareDate"], reverse=True)
    return runs[:limit]


# ---------- Insights API stubs (match your JS expectations) ------------------
# These should read/write JSON files under HISTORY_FOLDER.
# For now, you can leave your existing implementations here and just
# update them later to use the new comparison outputs.

@app.route("/api/history", methods=["GET"])
def api_history():
    """
    Return a list of available JSON snapshots for the given domain.

    Looks in RESULT_FOLDER for files like:
      analysis_summary_<domain>_YYYYMMDD_HHMMSS.json
    and exposes light metadata used by the Insights UI.
    """
    domain = (request.args.get("domain") or "").lower()
    if domain not in ("apm", "brum", "mrum"):
        return jsonify({"error": "Invalid domain."}), 400

    folder = RESULT_FOLDER
    prefix = f"analysis_summary_{domain}_"   # <-- matches your filenames

    items = []

    if not os.path.isdir(folder):
        return jsonify({"domain": domain.upper(), "items": []})

    for name in sorted(os.listdir(folder), reverse=True):
        if not (name.startswith(prefix) and name.endswith(".json")):
            continue

        path = os.path.join(folder, name)
        meta = {}
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            meta = payload.get("meta") or {}
        except Exception:
            meta = {}

        items.append(
            {
                "file": name,
                "timestamp": meta.get("compareDate", ""),
                "controller": meta.get("controller"),
                "prev": meta.get("previousDate"),
                "curr": meta.get("currentDate"),
            }
        )

    # optional controller filter
    controller_q = request.args.get("controller")
    if controller_q:
        want = _slug(controller_q)
        filtered = []
        for it in items:
            if it["controller"] and _slug(it["controller"]) == want:
                filtered.append(it)
        items = filtered

    items.sort(key=lambda x: x["timestamp"] or "", reverse=True)
    return jsonify({"domain": domain.upper(), "items": items})


@app.route("/api/apps", methods=["GET"])
def api_apps():
    """
    Return list of application names for a given domain & snapshot.

    If ?file=<name> is not provided, uses the latest snapshot for that domain.
    """
    domain = (request.args.get("domain") or "APM").upper()
    folder = RESULT_FOLDER

    # Optional explicit file selection
    file_name = request.args.get("file")

    def _latest_file_for_domain() -> Optional[str]:
        prefix = f"analysis_summary_{domain.lower()}_"
        if not os.path.isdir(folder):
            return None
        candidates = [
            n
            for n in os.listdir(folder)
            if n.startswith(prefix) and n.endswith(".json")
        ]
        if not candidates:
            return None
        # Files are timestamped; sorted() gives oldest->newest; we want newest
        return sorted(candidates)[-1]

    if not file_name:
        file_name = _latest_file_for_domain()

    if not file_name:
        # No snapshots yet for this domain
        return jsonify({"apps": []})

    path = os.path.join(folder, file_name)
    if not os.path.exists(path):
        return jsonify({"apps": []})

    try:
        with open(path, "r", encoding="utf-8") as f:
            payload = json.load(f)
        apps = payload.get("apps", {}).get("names", []) or []
    except Exception:
        apps = []

    return jsonify({"apps": apps})


@app.route("/api/insights", methods=["GET"])
def api_insights():
    domain = (request.args.get("domain") or "").upper()
    app_name = request.args.get("app") or ""
    file = request.args.get("file") or ""  # optional: specific summary filename

    if domain not in ("APM", "BRUM", "MRUM") or not app_name:
        return jsonify({"error": "Missing domain or app."}), 400

    folder = RESULT_FOLDER

    # Choose file: specific or latest for domain.
    if file:
        path = os.path.join(folder, file)
    else:
        prefix = f"analysis_summary_{domain.lower()}_"
        try:
            names = [
                n for n in os.listdir(folder)
                if n.startswith(prefix) and n.endswith(".json")
            ]
        except FileNotFoundError:
            names = []
        names.sort(reverse=True)  # newest first
        path = os.path.join(folder, names[0]) if names else ""

    if not path or not os.path.exists(path):
        return jsonify({"error": "Snapshot not found."}), 404

    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    apps_index = (data.get("appsIndex") or {})
    entry = apps_index.get(app_name)
    if not entry:
        # Try a normalized match
        key = app_name.strip().lower()
        for k, v in apps_index.items():
            if k.strip().lower() == key:
                entry = v
                break

    if not entry:
        return jsonify({"error": "App not found in snapshot."}), 404

    areas = entry.get("areas", [])
    detail = entry.get("detail", {})

    return jsonify(
        {
            "domain": domain,
            "app": app_name,
            "areas": areas,
            "detail": detail,
            "meta": data.get("meta", {}),
        }
    )



@app.route("/api/trends/runs", methods=["GET"])
def api_trends_runs():
    domain = (request.args.get("domain") or "").lower()
    if domain not in ("apm", "brum", "mrum"):
        return jsonify({"error": "Invalid domain."}), 400

    controller = request.args.get("controller")
    try:
        limit = int(request.args.get("limit", "20"))
    except ValueError:
        limit = 20

    baseline = (request.args.get("baseline") or "").lower()

    # 🔹 Use the module-level RESULT_FOLDER (same as other APIs)
    folder = RESULT_FOLDER

    runs = scan_runs(folder, domain=domain, controller_filter=controller, limit=limit)

    if baseline == "earliestprev":
        prevs = [r["sortPrev"] for r in runs if r.get("sortPrev")]
        if prevs:
            earliest_prev = min(prevs)
            runs = [r for r in runs if r.get("sortPrev") == earliest_prev]

    series = [
        {
            "compareDate": r["compareDate"],
            "previousDate": r["previousDate"],
            "currentDate": r["currentDate"],
            "improved": r["improved"],
            "degraded": r["degraded"],
            "percentage": r["percentage"],
            "tiers": r["tiers"],
            "file": r["file"],
        }
        for r in runs
    ]

    label = controller or (runs[0]["controller"] if runs else None)
    return jsonify(
        {
            "domain": domain.upper(),
            "controller": label,
            "count": len(series),
            "items": series,
        }
    )

if __name__ == "__main__":
    def open_browser():
        webbrowser.open("http://127.0.0.1:5000")

    # Give Flask a moment to start, then open the browser
    threading.Timer(1.0, open_browser).start()
    app.run(host="127.0.0.1", port=5000, debug=False)


