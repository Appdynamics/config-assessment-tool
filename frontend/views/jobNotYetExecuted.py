import os
import json
from urllib import parse

import requests
import streamlit as st
from docker import APIClient
from FileHandler import openFile, openFolder
from utils.docker_utils import isDocker, runConfigAssessmentTool


def jobNotYetExecuted(client: APIClient, jobName: str, debug: bool, concurrentConnections: int, platformStr: str, tag: str):
    st.header(f"{jobName}")
    (
        openJobFileColumn,
        openThresholdsFileColumn,
        _,
    ) = st.columns([1, 2, 2])

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

    thresholdsColumn, warningColumn, runColumn = st.columns([1, 1, 0.3])
    thresholdsFiles = [f[: len(f) - 5] for f in os.listdir("../input/thresholds")]

    if jobName in thresholdsFiles:
        thresholdsFiles.remove(jobName)
        thresholdsFiles.insert(0, jobName)
    elif "DefaultThresholds" in thresholdsFiles:
        thresholdsFiles.remove("DefaultThresholds")
        thresholdsFiles.insert(0, "DefaultThresholds")

    thresholds = thresholdsColumn.selectbox("Specify Thresholds File", thresholdsFiles, key=f"{jobName}-new")

    warningColumn.text("")
    warningColumn.warning(f"Job has not yet been run")

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

    dynamicCredentials = st.expander("Pass credentials dynamically (optional)")
    dynamicCredentials.write("Credentials will be changed for all jobs in the job file.")

    # Define columns for username, password, and authentication type
    usrNameCol, pwdCol, authTypeCol, dynChckCol = dynamicCredentials.columns(4)

    # Dropdown for authentication type
    authType = authTypeCol.selectbox(
        label="Auth Type",
        options=["basic", "secret", "token"],
        key=f"JobFile:{jobName}-authType"
    )

    # Dynamically set labels for username and password based on the selected authentication type
    if authType == "token":
        username_label = "Client Name"
        password_label = "Temporary Access Token"
    elif authType == "secret":
        username_label = "Client Name"
        password_label = "Client Secret"
    else:  # Default to "basic auth"
        username_label = "New Username"
        password_label = "New Password"

    # Input field for new username
    newUsrName = usrNameCol.text_input(
        label=username_label,
        value="AzureDiamond",
        key=f"JobFile:{jobName}-usrCol"
    )

    # Input field for new password
    newPwd = pwdCol.text_input(
        label=password_label,
        value="hunter2",
        type="password",
        key=f"JobFile:{jobName}-pwdCol"
    )

    # Dynamic credentials checkbox
    dynChckCol.text("")
    dynChckCol.text("")
    dynamicCheck = dynChckCol.checkbox("Dynamic Credentials", key=f"JobFile:{jobName}-chckCol")

    # Logic for running the job
    runColumn.text("")  # vertical padding
    if runColumn.button(f"Run", key=f"JobFile:{jobName}-Thresholds:{thresholds}-JobType:extract"):
        username = newUsrName if dynamicCheck else None
        password = newPwd if dynamicCheck else None
        auth_method = authType if dynamicCheck else None
        runConfigAssessmentTool(client, jobName, thresholds, debug,
                                concurrentConnections, username, password,
                                auth_method, platformStr, tag)
