import json
import os
import time
from typing import Any, Tuple
from urllib import parse

import requests
import streamlit as st
from FileHandler import openFolder
from utils.docker_utils import isDocker
from utils.stdlib_utils import base64Encode
from utils.streamlit_utils import rerun


def header() -> tuple[bool, bool]:
    st.set_page_config(
        page_title="config-assessment-tool",
    )
    st.markdown(
        """
        <style>
            h1 {
                text-align: center;
            }
            .stTextArea textarea { 
                font-family: monospace;
                font-size: 15px; 
            }
            .block-container {
                min-width: 1000px;
            }
            .info-bubble {
                display: inline-block;
                background-color: #e0e0e0;
                color: #333;
                border-radius: 12px;
                padding: 2px 10px;
                font-size: 12px;
                margin-left: 8px;
                vertical-align: middle;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("config-assessment-tool")
    st.markdown("---")

    if st.button("Open Jobs Folder"):
        if not isDocker():
            openFolder("../input/jobs")
        else:
            payload = {"type": "folder", "path": "input/jobs"}
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:16225?{payload}")

    if st.button("Open Thresholds Folder"):
        if not isDocker():
            openFolder("../input/thresholds")
        else:
            payload = {"type": "folder", "path": "input/thresholds"}
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:16225?{payload}")

    open_archive = st.button("Open Archive Folder")
    if open_archive:
        if not isDocker():
            openFolder("../output/archive")
        else:
            payload = {"type": "folder", "path": "output/archive"}
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:16225?{payload}")

    # Example: allow downloading a sample report file from archive folder
    sample_file_path = "../output/archive/sample_report.json"
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

        hostCol, portCol, _ = st.columns(3)
        host = hostCol.text_input(label="host", value="acme.saas.appdynamics.com")
        port = portCol.number_input(label="port", value=443)

        accountCol, usernameCol, pwdCol = st.columns(3)
        account = accountCol.text_input(label="account", value="acme")
        username = usernameCol.text_input(label="username", value="foo")
        pwd = pwdCol.text_input(label="password", value="hunter2", type="password")

        if st.form_submit_button("create"):
            job_file_path = f"../input/jobs/{host[:host.index('.')]}.json"
            with open(job_file_path, "w", encoding="ISO-8859-1") as f:
                json.dump(
                    [
                        {
                            "host": host,
                            "port": port,
                            "ssl": True,
                            "account": account,
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
                st.info(f"Successfully created job '{host[:host.index('.')]}'")
            else:
                st.error(f"Failed to create job '{host[:host.index('.')]}'")

            time.sleep(2)
            rerun()

    optionsCol1, optionsCol2, _ = st.columns(3)
    debug = optionsCol1.checkbox("Enable Debug")
    throttleNetworkConnections = optionsCol2.checkbox("Throttle Network Connections")

    st.markdown("---")

    return debug, throttleNetworkConnections