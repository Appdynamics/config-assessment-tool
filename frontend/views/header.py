import json
import os
import time
from typing import Tuple, Any
from urllib import parse

import requests
import streamlit as st

from FileHandler import openFolder
from utils.docker_utils import isDocker


def header() -> tuple[bool, bool]:
    st.markdown(
        f"""
                <style>
                    h1{{
                        text-align: center;
                    }}
                    .reportview-container .main .block-container{{
                        max-width: {1000}px;
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
        pwd = pwdCol.text_input(label="password", value="hunter1", type="password")

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

            # if file exists
            if os.path.exists("../input/thresholds/DefaultThresholds.json"):
                defaultThresholds = json.loads(open("../input/thresholds/DefaultThresholds.json").read())
                with open(f"../input/thresholds/{host[:host.index('.')]}.json", "w", encoding="ISO-8859-1") as f:
                    json.dump(defaultThresholds, fp=f, ensure_ascii=False, indent=4)
                if os.path.exists(f"../input/thresholds/{host[:host.index('.')]}.json"):
                    st.info(f"Successfully created thresholds for job '{host[:host.index('.')]}'")
                else:
                    st.error(f"Failed to create thresholds for job '{host[:host.index('.')]}'")
            else:
                st.error("Failed to create thresholds for job, DefaultThresholds.json not found")

            # small delay to see job ended
            time.sleep(2)
            # refresh the page to see newly generated report
            raise st.script_runner.RerunException(st.script_request_queue.RerunData(None))

    debug = optionsColumn.checkbox("Enable Debug")
    throttleNetworkConnections = optionsColumn.checkbox("Throttle Network Connections")
    st.markdown("""---""")

    return debug, throttleNetworkConnections
