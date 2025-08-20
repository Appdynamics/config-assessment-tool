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

    # Initialize run_with_car flag
    run_with_car = False

    # Define columns for the main button row: Open JobFile, Open Thresholds File, Open Output Folder
    # Adjust column weights as needed for desired spacing
    col_job_file, col_thresholds_file, col_output_folder = st.columns([1, 1, 1])

    # Place "Open JobFile" button
    col_job_file.text("")  # vertical padding
    col_job_file.text("")  # vertical padding
    if col_job_file.button(f"Open JobFile", key=f"{jobName}-jobfile"):
        file_path = f"input/jobs/{jobName}.json" if isDocker() else f"../input/jobs/{jobName}.json"
        if os.path.exists(file_path):
            with st.expander(f"📂 {jobName}.json", expanded=True):
                with open(file_path) as f:
                    data = json.load(f)
                    st.json(data)
        else:
            st.warning(f"File not found: {file_path}")

    # Place "Open Thresholds File" button (always present)
    col_thresholds_file.text("")  # vertical padding
    col_thresholds_file.text("")  # vertical padding
    # The actual button logic will be inside job_executed/else blocks, but the button itself is placed here.

    # Place "Open Output Folder" button (conditionally present)
    if job_executed:
        col_output_folder.text("")  # vertical padding
        col_output_folder.text("")  # vertical padding
        if col_output_folder.button(f"Open Output Folder", key=f"{jobName}-outputFolder"):
            if not isDocker():
                openFolder(f"../output/{jobName}")
            else:
                payload = {"type": "folder", "path": f"output/{jobName}"}
                payload = parse.urlencode(payload)
                requests.get(f"http://host.docker.internal:16225?{payload}")

    # Define a separate column for the "Run with CAR" checkbox on the next row
    col_run_with_car = st.columns([1])[0] # Use a single column for the checkbox

    # Place the "Run with CAR" checkbox here. Its value will be used in the runConfigAssessmentTool call.
    col_run_with_car.text("") # vertical padding for the checkbox row
    run_with_car = col_run_with_car.checkbox("Generate ConfigurationAnalysisReport(CAR).xlsx report used mainly for Professional Services teams", key=f"{jobName}-run-with-car-checkbox")

    # --- START: Moved Dynamic credentials section here ---
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
    # --- END: Moved Dynamic credentials section here ---


    # If job was previously executed, add last run info and thresholds selectbox
    if job_executed:
        # Load info.json for last run and thresholds
        info = json.loads(open(info_path).read())

        thresholdsFiles = [f[: -5] for f in os.listdir("../input/thresholds")]
        if info["thresholds"] in thresholdsFiles:
            default_idx = thresholdsFiles.index(info["thresholds"])
        else:
            default_idx = 0

        # Row for Thresholds Selectbox, Last Run Info, and Run button
        thresholdsColumn, lastRunColumn, runColumn = st.columns([1, 1, 0.3])

        lastRunColumn.text("")  # vertical padding
        lastRunColumn.info(f'Last Run: {datetime.fromtimestamp(info["lastRun"], get_localzone()).strftime("%m-%d-%Y at %H:%M:%S")}')

        thresholds = thresholdsColumn.selectbox(
            "Specify Thresholds File",
            thresholdsFiles,
            index=default_idx,
            key=f"{jobName}-new",
        )

        # Logic for "Open Thresholds File" button (using col_thresholds_file defined above)
        if col_thresholds_file.button(f"Open Thresholds File", key=f"{jobName}-thresholds"):
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

        # Run button logic for executed job
        runColumn.text("")  # vertical padding
        if runColumn.button(f"Run", key=f"JobFile:{jobName}-Thresholds:{thresholds}-JobType:extract"):
            username = newUsrName if dynamicCheck else None
            password = newPwd if dynamicCheck else None
            auth_method = authType if dynamicCheck else None
            runConfigAssessmentTool(
                client, jobName, thresholds, debug,
                concurrentConnections, username, password,
                auth_method, platformStr, tag,
                run_with_car # Pass the new parameter
            )

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

        # Row for Thresholds Selectbox, Warning, and Run button
        thresholdsColumn, warningColumn, runColumn = st.columns([1, 1, 0.3])

        warningColumn.text("")
        warningColumn.warning(f"Job has not yet been run")

        thresholds = thresholdsColumn.selectbox("Specify Thresholds File", thresholdsFiles, key=f"{jobName}-new")

        # Logic for "Open Thresholds File" button (using col_thresholds_file defined above)
        if col_thresholds_file.button(f"Open Thresholds File", key=f"{jobName}-thresholds"):
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

        # Run button logic for non-executed job
        runColumn.text("")  # vertical padding
        if runColumn.button(f"Run", key=f"JobFile:{jobName}-Thresholds:{thresholds}-JobType:extract"):
            username = newUsrName if dynamicCheck else None
            password = newPwd if dynamicCheck else None
            auth_method = authType if dynamicCheck else None
            runConfigAssessmentTool(
                client, jobName, thresholds, debug,
                concurrentConnections, username, password,
                auth_method, platformStr, tag,
                run_with_car # Pass the new parameter
            )