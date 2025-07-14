import os
import json
from datetime import datetime
from urllib import parse

import requests
import streamlit as st
from docker import APIClient
from tzlocal import get_localzone
from utils.docker_utils import isDocker, runConfigAssessmentTool
from FileHandler import openFolder


def jobHandler(client: APIClient, jobName: str, debug: bool, concurrentConnections: int, platformStr: str, tag: str):
    info_path = f"../output/{jobName}/info.json"
    job_executed = os.path.exists(info_path)

    st.header(f"{jobName}")

    # Common UI elements for opening job file
    openJobFileColumn, openThresholdsFileColumn, *rest = st.columns([1, 1.3, 1])

    openJobFileColumn.text("")  # vertical padding
    openJobFileColumn.text("")  # vertical padding

    if openJobFileColumn.button(f"Open JobFile", key=f"{jobName}-jobfile"):
        file_path = f"input/jobs/{jobName}.json" if isDocker() else f"../input/jobs/{jobName}.json"
        if os.path.exists(file_path):
            with st.expander(f"📂 {jobName}.json", expanded=True):
                with open(file_path) as f:
                    data = json.load(f)
                    st.json(data)
        else:
            st.warning(f"File not found: {file_path}")

    # If job was previously executed, add output folder button and last run info
    if job_executed:
        # Open Output Folder button
        openOutputFolderColumn = st.columns([1.25])[0]
        openOutputFolderColumn.text("")  # vertical padding
        openOutputFolderColumn.text("")  # vertical padding
        if openOutputFolderColumn.button(f"Open Output Folder", key=f"{jobName}-outputFolder"):
            if not isDocker():
                openFolder(f"../output/{jobName}")
            else:
                payload = {"type": "folder", "path": f"output/{jobName}"}
                payload = parse.urlencode(payload)
                requests.get(f"http://host.docker.internal:16225?{payload}")

        # Load info.json for last run and thresholds
        info = json.loads(open(info_path).read())

        thresholdsFiles = [f[: -5] for f in os.listdir("../input/thresholds")]
        if info["thresholds"] in thresholdsFiles:
            default_idx = thresholdsFiles.index(info["thresholds"])
        else:
            default_idx = 0

        thresholdsColumn, lastRunColumn, runColumn = st.columns([1, 1, 0.3])

        lastRunColumn.text("")  # vertical padding
        lastRunColumn.info(f'Last Run: {datetime.fromtimestamp(info["lastRun"], get_localzone()).strftime("%m-%d-%Y at %H:%M:%S")}')

        thresholds = thresholdsColumn.selectbox(
            "Specify Thresholds File",
            thresholdsFiles,
            index=default_idx,
            key=f"{jobName}-new",
        )

        openThresholdsFileColumn.text("")  # vertical padding
        openThresholdsFileColumn.text("")  # vertical padding

        if openThresholdsFileColumn.button(f"Open Thresholds File", key=f"{jobName}-thresholds"):
            file_path = f"input/thresholds/{thresholds}.json" if isDocker() else f"../input/thresholds/{thresholds}.json"
            if os.path.exists(file_path):
                with st.expander(f"📂 {thresholds}.json", expanded=True):
                    with open(file_path) as f:
                        data = json.load(f)
                        formatted = json.dumps(data, indent=2)
                        st.markdown(
                            f"""
                            <div style="max-height: 240px; overflow-y: scroll; border: 1px solid #ccc; padding: 8px; background-color: #f9f9f9;">
                            <pre>{formatted}</pre>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
            else:
                st.warning(f"File not found: {file_path}")

    else:
        # Job not yet executed
        thresholdsFiles = [f[: -5] for f in os.listdir("../input/thresholds")]

        # Prioritize jobName or DefaultThresholds at top of list
        if jobName in thresholdsFiles:
            thresholdsFiles.remove(jobName)
            thresholdsFiles.insert(0, jobName)
        elif "DefaultThresholds" in thresholdsFiles:
            thresholdsFiles.remove("DefaultThresholds")
            thresholdsFiles.insert(0, "DefaultThresholds")

        thresholdsColumn, warningColumn, runColumn = st.columns([1, 1, 0.3])

        warningColumn.text("")
        warningColumn.warning(f"Job has not yet been run")

        thresholds = thresholdsColumn.selectbox("Specify Thresholds File", thresholdsFiles, key=f"{jobName}-new")

        openThresholdsFileColumn.text("")  # vertical padding
        openThresholdsFileColumn.text("")  # vertical padding

        if openThresholdsFileColumn.button(f"Open Thresholds File", key=f"{jobName}-thresholds"):
            file_path = f"input/thresholds/{thresholds}.json" if isDocker() else f"../input/thresholds/{thresholds}.json"
            if os.path.exists(file_path):
                with st.expander(f"📂 {thresholds}.json", expanded=True):
                    with open(file_path) as f:
                        data = json.load(f)
                        formatted = json.dumps(data, indent=2)
                        st.markdown(
                            f"""
                            <div style="max-height: 240px; overflow-y: scroll; border: 1px solid #ccc; padding: 8px; background-color: #f9f9f9;">
                            <pre>{formatted}</pre>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
            else:
                st.warning(f"File not found: {file_path}")

    # Dynamic credentials section (common to both cases)
    dynamicCredentials = st.expander("Pass credentials dynamically (optional)")
    dynamicCredentials.write("Credentials will be changed for all jobs in the job file.")

    usrNameCol, pwdCol, authTypeCol, dynChckCol = dynamicCredentials.columns(4)

    authType = authTypeCol.selectbox(
        label="Auth Type",
        options=["basic", "secret", "token"],
        key=f"JobFile:{jobName}-authType"
    )

    if authType == "token":
        if job_executed:
            username_label = "API Client Username"
            password_label = "API Client Token"
        else:
            username_label = "Client Name"
            password_label = "Temporary Access Token"
    elif authType == "secret":
        username_label = "Client ID" if job_executed else "Client Name"
        password_label = "Client Secret"
    else:  # basic auth
        username_label = "New Username"
        password_label = "New Password"

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

    # Run button logic (common)
    runColumn = None
    if job_executed:
        # runColumn was created above for executed job
        runColumn = st.columns([0.3])[0]
    else:
        # runColumn was created above for not executed job
        runColumn = st.columns([0.3])[0]

    runColumn.text("")  # vertical padding
    if runColumn.button(f"Run", key=f"JobFile:{jobName}-Thresholds:{thresholds}-JobType:extract"):
        username = newUsrName if dynamicCheck else None
        password = newPwd if dynamicCheck else None
        auth_method = authType if dynamicCheck else None
        runConfigAssessmentTool(
            client, jobName, thresholds, debug,
            concurrentConnections, username, password,
            auth_method, platformStr, tag
        )