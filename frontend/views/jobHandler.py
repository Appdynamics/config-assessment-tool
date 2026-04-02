import os
import json
import time
import asyncio
import logging
import multiprocessing
import platform
import subprocess
from datetime import datetime

import requests
import streamlit as st
from tzlocal import get_localzone
from streamlit_modal import Modal
import streamlit.components.v1 as components

from utils.streamlit_utils import rerun

# --- Helper Functions ---

def run_backend_process(job_details):
    """Target function for the separate process to run the backend engine."""

    # Ensure sys.path includes the backend directory when running in a subprocess within the frozen bundle.
    # This mirrors the logic in bundle_main.py to ensure top-level imports like 'api' work.
    import sys
    import os
    if getattr(sys, 'frozen', False):
        backend_path = os.path.join(sys._MEIPASS, 'backend')
        if backend_path not in sys.path:
            sys.path.append(backend_path)

    # This function runs in a completely separate process.
    # We need to re-import and initialize everything it needs.
    from backend.core.Engine import Engine
    from backend.util.logging_utils import initLogging

    # Initialize logging for this process. It will write to the same log file.
    initLogging(debug=job_details.get("debug", False))

    async def run_main():
        engine = Engine(
            job_details["job_file"],
            job_details["thresholds_file"],
            job_details["concurrent_connections"],
            job_details["username"],
            job_details["password"],
            job_details["auth_method"]
        )
        await engine.run()

    # Run the async engine code.
    asyncio.run(run_main())


def is_running_in_container():
    """
    Checks if the code is running inside a container by inspecting
    the environment. This works for Docker, containerd, and other
    runtimes that use standard cgroup paths.
    """
    # Check for a common file created by Docker.
    if os.path.exists('/.dockerenv'):
        return True

    # Check the cgroup of the init process for container-specific keywords.
    try:
        with open('/proc/1/cgroup', 'rt') as f:
            cgroup_content = f.read()
            if 'docker' in cgroup_content or 'kubepods' in cgroup_content:
                return True
    except FileNotFoundError:
        # /proc/1/cgroup does not exist on non-Linux systems.
        pass

    return False

def get_file_path(base, name):
    return f"input/{base}/{name}.json"

def handle_open_jobfile(file_path, title):
    # This function now only displays the file content
    if os.path.exists(file_path):
        with st.expander(f"📂 {title}", expanded=True):
            with open(file_path) as f:
                st.json(json.load(f))
    else:
        st.warning(f"File not found: {file_path}")

def show_thresholds_file(thresholds):
    file_path = get_file_path("thresholds", thresholds)
    if os.path.exists(file_path):
        with st.expander(f"📂 {thresholds}.json", expanded=True):
            with open(file_path) as f:
                st.json(json.load(f))
    else:
        st.warning(f"File not found: {file_path}")

def open_output_folder(jobName):
    """Opens a folder."""
    relative_path = f"output/{jobName}"

    if is_running_in_container():
        # Dynamically determine the host based on the environment
        # 'host.docker.internal' is a special DNS name that resolves to the host's IP.
        default_host = "host.docker.internal"

        # Allow overriding with an environment variable for flexibility
        file_handler_host = os.getenv("FILE_HANDLER_HOST", default_host)

        url = f"http://{file_handler_host}:16225/open_folder"
        try:
            response = requests.post(url, json={"path": relative_path}, timeout=5)
            response.raise_for_status()
            logging.info(f"Successfully requested to open folder: {relative_path}")
        except requests.exceptions.RequestException as e:
            st.error(f"Could not connect to FileHandler server at {url}. Is it running?")
            logging.error(f"Failed to open folder via FileHandler server: {e}")
    else:
        # Running locally, open directly
        logging.info("Running from source, opening folder directly: %s", relative_path)
        try:
            abs_path = os.path.abspath(relative_path)
            if not os.path.isdir(abs_path):
                st.error(f"Directory does not exist: {abs_path}")
                return

            system = platform.system()
            if system == "Windows":
                os.startfile(abs_path)
            elif system == "Darwin":  # macOS
                subprocess.run(["open", abs_path], check=True)
            else:  # Linux
                # Check for DISPLAY environment variable to see if UI is available
                if not os.environ.get('DISPLAY'):
                    st.warning(f"Cannot open folder '{abs_path}' directly: Running in a headless environment (no display detected).")
                    logging.warning(f"Cannot open folder '{abs_path}' directly: Running in a headless environment.")
                    return
                subprocess.run(["xdg-open", abs_path], check=True)
        except Exception as e:
            st.error(f"Failed to open folder directly: {e}")
            logging.error(f"Failed to open folder directly: {e}")

def dynamic_credentials_section(job_executed, jobName):
    dynamicCredentials = st.expander("Pass credentials dynamically (optional)")
    dynamicCredentials.write("Credentials will be changed for all jobs in the job file.")
    usrNameCol, pwdCol, authTypeCol, dynChckCol = dynamicCredentials.columns(4)
    authType = authTypeCol.selectbox(
        label="Auth Type",
        options=["basic", "secret", "token"],
        key=f"JobFile:{jobName}-authType"
    )
    labels = {
        "token": ("API Client Username", "API Client Token"),
        "secret": ("Client ID", "Client Secret"),
        "basic": ("New Username", "New Password")
    }
    username_label, password_label = labels.get(authType, labels["basic"])
    newUsrName = usrNameCol.text_input(
        label=username_label,
        value="AzureDiamond",
        key=f"JobFile:{jobName}-usrCol"
    )
    newPwd = pwdCol.text_input(
        label=password_label,
        value="hunter2",
        type="password",
        key=f"JobFile:{jobName}-pwdCol"
    )
    dynChckCol.text("")
    dynChckCol.text("")
    dynamicCheck = dynChckCol.checkbox("Dynamic Credentials", key=f"JobFile:{jobName}-chckCol")
    return newUsrName, newPwd, authType, dynamicCheck

def handle_run(runColumn, jobName, thresholds, debug, concurrentConnections, newUsrName, newPwd, authType, dynamicCheck, log_modal):
    runColumn.text("")
    if runColumn.button(f"Run", key=f"JobFile:{jobName}-Thresholds:{thresholds}-JobType:extract"):
        st.session_state.job_to_run_details = {
            "job_file": jobName,
            "thresholds_file": thresholds,
            "concurrent_connections": concurrentConnections,
            "username": newUsrName if dynamicCheck else None,
            "password": newPwd if dynamicCheck else None,
            "auth_method": authType if dynamicCheck else None,
            "debug": debug
        }
        log_modal.open()
        rerun()

def tail_file(filepath, n_lines=50):
    """Reads the last N lines from a file."""
    try:
        with open(filepath, "r") as f:
            lines = f.readlines()
            return "".join(lines[-n_lines:])
    except FileNotFoundError:
        return "Log file not found."
    except Exception as e:
        return f"Error reading log file: {e}"

# --- Main Component ---

def jobHandler(jobName: str, debug: bool, concurrentConnections: int):
    st.header(f"{jobName}")

    col_job_file, col_thresholds_file, col_output_folder = st.columns([1, 1, 1])

    # Column 1: Job File
    col_job_file.text("")
    col_job_file.text("")
    if col_job_file.button(f"Open JobFile", key=f"{jobName}-jobfile"):
        handle_open_jobfile(f"input/jobs/{jobName}.json", f"{jobName}.json")

    # Column 2: Thresholds File
    col_thresholds_file.text("")
    col_thresholds_file.text("")

    # Column 3: Output Folder
    info_path = f"output/{jobName}/info.json"
    job_executed = os.path.exists(info_path)
    if job_executed:
        col_output_folder.text("")
        col_output_folder.text("")
        if col_output_folder.button(f"Open Output Folder", key=f"{jobName}-outputFolder"):
            open_output_folder(jobName)

    # Dynamic Credentials
    newUsrName, newPwd, authType, dynamicCheck = dynamic_credentials_section(job_executed, jobName)

    # Thresholds Selection
    thresholds_dir = "input/thresholds"
    thresholdsFiles = []
    if os.path.exists(thresholds_dir) and os.path.isdir(thresholds_dir):
        thresholdsFiles = [f[:-5] for f in os.listdir(thresholds_dir) if f.endswith('.json')]

    if jobName in thresholdsFiles:
        thresholdsFiles.remove(jobName)
        thresholdsFiles.insert(0, jobName)
    elif "DefaultThresholds" in thresholdsFiles:
        thresholdsFiles.remove("DefaultThresholds")
        thresholdsFiles.insert(0, "DefaultThresholds")

    # Main Action Row
    thresholdsColumn, infoColumn, runColumn = st.columns([1, 1, 0.3])

    if job_executed:
        try:
            with open(info_path) as f:
                info = json.load(f)
            last_run_str = datetime.fromtimestamp(info["lastRun"], get_localzone()).strftime("%m-%d-%Y at %H:%M:%S")
            infoColumn.text("")
            infoColumn.info(f'Last Run: {last_run_str}')
        except (IOError, json.JSONDecodeError, KeyError):
            infoColumn.text("")
            infoColumn.warning("Job has not yet been run or info file is invalid.")
    else:
        infoColumn.text("")
        infoColumn.warning("Job has not yet been run")

    # Log viewer Modal
    modal_title = "Log output for config-assessment-tool"

    # Check if job represents a running process to determine title state
    pid_peek = st.session_state.get(f"process_{jobName}")
    is_running_peek = False
    if pid_peek:
        try:
            os.kill(pid_peek, 0)
            is_running_peek = True
        except (OSError, SystemError):
            is_running_peek = False

    if not is_running_peek:
        # Check if the logs show completion
        log_peek = tail_file("logs/config-assessment-tool.log", 50)
        if "----------Complete----------" in log_peek:
            modal_title = "Log output for config-assessment-tool. JOB FINISHED! You may close the window"

    log_modal = Modal(modal_title, key=f"logs-modal-{jobName}", max_width=5000)

    # CSS to force the modal to 90% of screen width
    st.markdown("""
        <style>
        div[data-modal-container='true'] > div:first-child {
            width: 90vw !important;
            max-width: 90vw !important;
        }
        </style>
    """, unsafe_allow_html=True)

    if st.button("Show Logs", key=f"show-logs-{jobName}"):
        log_modal.open()

    if thresholdsFiles:
        thresholds = thresholdsColumn.selectbox("Specify Thresholds File", thresholdsFiles, index=0, key=f"{jobName}-new")
        # Connect the button in col_thresholds_file to its action here
        if col_thresholds_file.button(f"Open Thresholds File", key=f"{jobName}-thresholds"):
            show_thresholds_file(thresholds)
        handle_run(runColumn, jobName, thresholds, debug, concurrentConnections, newUsrName, newPwd, authType, dynamicCheck, log_modal)
    else:
        thresholdsColumn.warning("No threshold files found in `input/thresholds`.")

    # This block now handles both displaying logs and running the job.
    if log_modal.is_open():
        with log_modal.container():
            # If a job was just triggered, start it in a new process.
            if "job_to_run_details" in st.session_state and st.session_state.job_to_run_details["job_file"] == jobName:
                details = st.session_state.pop("job_to_run_details")
                p = multiprocessing.Process(target=run_backend_process, args=(details,))
                p.start()
                st.session_state[f"process_{jobName}"] = p.pid
                logging.info(f"Started job '{jobName}' in process with PID {p.pid}")

            # Check if a process is running for this job
            is_running = False
            pid = st.session_state.get(f"process_{jobName}")
            if pid:
                # Check process status. Catches both standard errors and Windows-specific internal errors.
                try:
                    os.kill(pid, 0)
                    is_running = True
                except (OSError, SystemError):
                    # Process not found or inaccessible (WinError 1 / ESRCH)
                    is_running = False
                    del st.session_state[f"process_{jobName}"]

            log_file = "logs/config-assessment-tool.log"
            log_placeholder = st.empty()
            log_container_id = f"log-container-{jobName.replace(' ', '-')}"

            def display_logs(num_lines, auto_scroll):
                log_content = tail_file(log_file, num_lines)

                # Log container HTML
                log_html = f'<div id="{log_container_id}" style="height: 400px; overflow-y: scroll; overflow-x: auto; border: 1px solid #ccc; padding: 10px; background-color: #f0f2f6; font-family: monospace; white-space: pre; font-size: 7px !important; line-height: 1.2 !important;">{log_content}</div>'
                log_placeholder.markdown(log_html, unsafe_allow_html=True)

                if auto_scroll:
                    js_autoscroll = f"""
                        <script>
                            (function() {{
                                setTimeout(function() {{
                                    try {{
                                        const logContainer = window.parent.document.getElementById('{log_container_id}');
                                        if (logContainer) {{
                                            logContainer.scrollTop = logContainer.scrollHeight;
                                        }}
                                    }} catch (e) {{
                                        console.log("Could not scroll log container: " + e);
                                    }}
                                }}, 100);
                            }})();
                        </script>
                    """
                    components.html(js_autoscroll, height=0)
                return log_content

            # Live-tail the logs while the process is running
            if is_running:
                display_logs(500, auto_scroll=True)
                time.sleep(1) # Rerun every second to update logs
                st.rerun()
            else:
                # Display final log state after the job is finished or when just viewing logs
                log_content = display_logs(1000, auto_scroll=True)
                if "----------Complete----------" in log_content:
                    st.markdown(
                        """
                        <div style='text-align: center; margin-top: 10px;'>
                            <h2 style='color: green; margin-bottom: 0px;'>Job Finished!</h2>
                            <p style='font-size: 1.2em; color: #555;'>You may close this window.</p>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

