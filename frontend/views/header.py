import json
import logging
import os
import time
import platform
import subprocess

import requests
import streamlit as st

from utils.stdlib_utils import base64Encode
from utils.streamlit_utils import rerun


def is_running_in_container():
    """
    Checks if the code is running inside a container.
    """
    # Check for /.dockerenv file
    if os.path.exists('/.dockerenv'):
        return True
    # Check for cgroup info
    try:
        with open('/proc/1/cgroup', 'rt') as f:
            cgroup_content = f.read()
            return 'docker' in cgroup_content or 'kubepods' in cgroup_content
    except FileNotFoundError:
        # This check is not available on all OSes (e.g., macOS, Windows)
        pass
    # Check for container environment variable
    if os.environ.get('CONTAINER_RUNTIME'):
        return True
    return False


def get_filehandler_host():
    host = os.environ.get("FILEHANDLER_HOST")
    if host:
        return host
    if is_running_in_container():
        return "host.docker.internal"
    return "localhost"


def open_folder_via_service(path: str):
    if is_running_in_container():
        logging.info("Running in container, opening folder via service: %s", path)
        host = get_filehandler_host()
        try:
            response = requests.post(
                f"http://{host}:16225/open_folder",
                json={"path": path},
                headers={"Content-Type": "application/json"},
                timeout=10
            )
            if response.status_code != 200:
                st.error(f"Failed to open folder: {response.text}")
        except Exception as e:
            st.error(f"Error contacting FileHandler service: {e}")
    else:
        logging.info("Running from source, opening folder directly: %s", path)
        try:
            abs_path = os.path.abspath(path)
            if not os.path.exists(abs_path):
                try:
                    os.makedirs(abs_path)
                    logging.info(f"Created directory: {abs_path}")
                except OSError as e:
                    st.error(f"Directory does not exist and could not be created: {abs_path} ({e})")
                    return

            if not os.path.isdir(abs_path):
                st.error(f"Path exists but is not a directory: {abs_path}")
                return

            system = platform.system()
            if system == "Windows":
                os.startfile(abs_path)
            elif system == "Darwin":  # macOS
                subprocess.run(["open", abs_path], check=True)
            else:  # Linux
                subprocess.run(["xdg-open", abs_path], check=True)
        except Exception as e:
            st.error(f"Failed to open folder directly: {e}")


def header() -> tuple[bool, bool]:
    st.set_page_config(page_title="config-assessment-tool")

    top_col1, top_col2 = st.columns([8, 1])
    with top_col2:
        if st.button("Shutdown 🛑", key="global_shutdown_btn", help="Stop the application server"):
            st.markdown(
                """
                <div style="
                    position: fixed;
                    top: 0;
                    left: 0;
                    width: 100vw;
                    height: 100vh;
                    background-color: rgba(0, 0, 0, 0.85);
                    z-index: 1000000;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                ">
                    <div style="
                        background-color: #ffffff;
                        padding: 40px;
                        border-radius: 12px;
                        text-align: center;
                        box-shadow: 0 10px 25px rgba(0,0,0,0.5);
                    ">
                        <h2 style="color: #333; margin: 0 0 15px 0;">Configuration Assessment Tool has been shut down.</h2>
                        <h4 style="color: #555; margin: 0; font-weight: normal;">You may close this tab.</h4>
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )
            time.sleep(3)
            os._exit(0)

    st.markdown(
        """
        <style>
            h1 { text-align: center; }
            .stTextArea textarea { font-family: monospace; font-size: 7px; }
            .block-container { min-width: 1000px; }
            .info-bubble {
                display: inline-block;
                background-color: #e0e0e0;
                color: #333;
                border-radius: 12px;
                padding: 2px 10px;
                font-size: 7px;
                margin-left: 8px;
                vertical-align: middle;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("config-assessment-tool")
    st.markdown("---")

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("Open Jobs Folder"):
            open_folder_via_service("input/jobs")
    with col2:
        if st.button("Open Thresholds Folder"):
            open_folder_via_service("input/thresholds")
    with col3:
        if st.button("Open Archive Folder"):
            open_folder_via_service("output/archive")


    sample_file_path = "output/archive/sample_report.json"
    if os.path.exists(sample_file_path):
        with open(sample_file_path, "rb") as f:
            file_bytes = f.read()
        st.download_button(
            label="Download Sample Report",
            data=file_bytes,
            file_name="sample_report.json",
            mime="application/json",
        )
    else:
        st.info("No sample report available for download.")

    st.markdown("---")

    newJobExpander = st.expander("Create New Job")
    with newJobExpander.form("NewJob"):
        st.write("Create new Job")

        hostCol, portCol, authTypeCol = st.columns(3)
        host = hostCol.text_input(label="host", value="acme.saas.appdynamics.com")
        port = portCol.number_input(label="port", value=443)
        authType = authTypeCol.selectbox("Auth Type", ["basic", "token", "secret"])

        accountCol, usernameCol, pwdCol = st.columns(3)
        account = accountCol.text_input(label="account", value="acme")
        username = usernameCol.text_input(label="username", value="foo")
        pwd = pwdCol.text_input(label="password", value="hunter2", type="password")

        if st.form_submit_button("create"):
            job_file_path = f"input/jobs/{host}.json"
            with open(job_file_path, "w", encoding="ISO-8859-1") as f:
                json.dump(
                    [
                        {
                            "host": host,
                            "port": port,
                            "ssl": True,
                            "account": account,
                            "authType": authType,
                            "username": username,
                            "pwd": base64Encode(f"CAT-ENCODED-{pwd}"),
                            "verifySsl": True,
                            "useProxy": True,
                            "applicationFilter": {"apm": ".*", "mrum": ".*", "brum": ".*"},
                            "timeRangeMins": 1440,
                        }
                    ],
                    fp=f,
                    ensure_ascii=False,
                    indent=4,
                )
            if os.path.exists(job_file_path):
                st.info(f"Successfully created job '{host}'")
            else:
                st.error(f"Failed to create job '{host}'")

            time.sleep(2)
            rerun()

    optionsCol1, optionsCol2, _ = st.columns(3)
    debug = optionsCol1.checkbox("Enable Debug")
    throttleNetworkConnections = optionsCol2.checkbox("Throttle Network Connections")

    st.markdown("---")

    return debug, throttleNetworkConnections