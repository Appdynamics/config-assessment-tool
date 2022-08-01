import json
import os
import time
from typing import Any, Tuple
from urllib import parse

import requests
import streamlit as st
from FileHandler import openFolder
from utils.docker_utils import isDocker
from utils.streamlit_utils import rerun


def header() -> tuple[bool, bool]:
    st.set_page_config(
        page_title="config-assessment-tool",
    )
    st.markdown(
        f"""
            <style>
                h1{{
                    text-align: center;
                }}
                .stTextArea textarea {{ 
                    font-family: monospace;
                    font-size: 15px; 
                }}
                .block-container{{
                    min-width: 1000px;
                }}
            </style>
            """,
        unsafe_allow_html=True,
    )

    st.title("config-assessment-tool")
    st.markdown("""---""")
    openJobsFolderColumn, openThresholdsFolderColumn, optionsColumn = st.columns(3)

    if openJobsFolderColumn.button(f"Open Jobs Folder"):
        if not isDocker():
            openFolder(f"../input/jobs")
        else:
            payload = {"type": "folder", "path": f"input/jobs"}
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:16225?{payload}")

    if openThresholdsFolderColumn.button(f"Open Thresholds Folder"):
        if not isDocker():
            openFolder(f"../input/thresholds")
        else:
            payload = {"type": "folder", "path": f"input/thresholds"}
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:16225?{payload}")
    newJobColumn, newThresholdColumn, _ = st.columns(3)

    newJob = st.expander("Create New Job")
    with newJob.form("NewJob"):
        st.write("Create new Job")

        hostCol, portCol, _ = st.columns(3)
        host = hostCol.text_input(label="host", value="acme.saas.appdynamics.com")
        port = portCol.number_input(label="port", value=443)

        accountCol, usernameCol, pwdCol = st.columns(3)
        account = accountCol.text_input(label="account", value="acme")
        username = usernameCol.text_input(label="username", value="foo")
        pwd = pwdCol.text_input(label="password", value="hunter2", type="password")

        if st.form_submit_button("create"):
            with open(f"../input/jobs/{host[:host.index('.')]}.json", "w", encoding="ISO-8859-1") as f:
                json.dump(
                    [
                        {
                            "host": host,
                            "port": port,
                            "ssl": True,
                            "account": account,
                            "username": username,
                            "pwd": pwd,
                            "verifySsl": True,
                            "useProxy": True,
                            "applicationFilter": {"apm": ".*", "mrum": ".*", "brum": ".*"},
                        }
                    ],
                    fp=f,
                    ensure_ascii=False,
                    indent=4,
                )
            if os.path.exists(f"../input/jobs/{host[:host.index('.')]}.json"):
                st.info(f"Successfully created job '{host[:host.index('.')]}'")
            else:
                st.error(f"Failed to create job '{host[:host.index('.')]}'")

            # small delay to see job ended
            time.sleep(2)
            # refresh the page to see newly generated report
            rerun()

    debug = optionsColumn.checkbox("Enable Debug")
    throttleNetworkConnections = optionsColumn.checkbox("Throttle Network Connections")
    st.markdown("""---""")

    return debug, throttleNetworkConnections
