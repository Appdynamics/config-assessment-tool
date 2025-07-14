import os
import json
import subprocess
import platform
from datetime import datetime
from urllib import parse

import requests
import streamlit as st
from docker import APIClient
from tzlocal import get_localzone
from utils.docker_utils import isDocker, runConfigAssessmentTool
from FileHandler import openFolder


def open_in_default_editor(file_path: str):
    if not os.path.exists(file_path):
        st.warning(f"File not found: {file_path}")
        return

    system_platform = platform.system()
    try:
        if system_platform == "Windows":
            os.startfile(file_path)
        else:
            subprocess.run(["open", file_path])
    except Exception as e:
        st.error(f"Failed to open file in default editor: {e}")


def load_json_file(file_path: str):
    if not os.path.exists(file_path):
        st.warning(f"File not found: {file_path}")
        return None
    try:
        with open(file_path, "r") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        st.error(f"Failed to parse JSON file: {e}")
        return None


def jobHandler(client: APIClient, jobName: str, debug: bool, concurrentConnections: int, platformStr: str, tag: str):
    info_path = f"../output/{jobName}/info.json"
    job_executed = os.path.exists(info_path)

    st.header(f"{jobName}")

    # Define columns for buttons depending on job execution status
    if job_executed:
        openEditCol, openOutputFolderCol, openThresholdsFileColumn = st.columns([1, 1.25, 1.3])
    else:
        openEditCol, openThresholdsFileColumn = st.columns([1, 1.3])
        openOutputFolderCol = None  # No output folder button if job not executed

    # Single button: View/Edit Job file
    openEditCol.text("")  # vertical padding
    openEditCol.text("")
    if openEditCol.button(f"View/Edit Job file", key=f"{jobName}-view-edit-jobfile"):
        st.session_state[f"{jobName}-editor-open"] = True

    # If editor is open, show editable text area with save button
    if st.session_state.get(f"{jobName}-editor-open", False):
        file_path = f"input/jobs/{jobName}.json" if isDocker() else f"../input/jobs/{jobName}.json"

        if os.path.exists(file_path):
            try:
                with open(file_path, "r") as f:
                    file_content = f.read()
            except Exception as e:
                st.error(f"Error reading file: {e}")
                file_content = ""
        else:
            st.warning(f"File not found: {file_path}")
            file_content = ""

        edited_text = st.text_area(f"Editing {jobName}.json", value=file_content, height=400, key=f"{jobName}-edit-textarea")

        if st.button("Save Changes", key=f"{jobName}-save-changes"):
            try:
                # Validate JSON before saving
                parsed_json = json.loads(edited_text)
                with open(file_path, "w") as f:
                    json.dump(parsed_json, f, indent=2)
                st.success("File saved successfully.")
                st.session_state[f"{jobName}-editor-open"] = False
            except json.JSONDecodeError as e:
                st.error(f"Invalid JSON format: {e}")

    # Open Output Folder button (only if job executed)
    if job_executed and openOutputFolderCol is not None:
        openOutputFolderCol.text("")  # vertical padding
        openOutputFolderCol.text("")  # vertical padding
        if openOutputFolderCol.button(f"Open Output Folder", key=f"{jobName}-outputFolder"):
            if not isDocker():
                openFolder(f"../output/{jobName}")
            else:
                payload = {"type": "folder", "path": f"output/{jobName}"}
                payload = parse.urlencode(payload)
                requests.get(f"http://host.docker.internal:16225?{payload}")

    if job_executed:
        # Load info.json for last run and thresholds
        info = None
        try:
            with open(info_path, "r") as f:
                info = json.load(f)
        except Exception as e:
            st.error(f"Failed to load info.json: {e}")

        thresholdsFiles = []
        try:
            thresholdsFiles = [f[:-5] for f in os.listdir("../input/thresholds")]
        except Exception as e:
            st.error(f"Failed to list thresholds files: {e}")

        if info and info.get("thresholds") in thresholdsFiles:
            default_idx = thresholdsFiles.index(info["thresholds"])
        else:
            default_idx = 0

        thresholdsColumn, lastRunColumn, runColumn = st.columns([1, 1, 0.3])

        lastRunColumn.text("")  # vertical padding
        if info and "lastRun" in info:
            lastRunColumn.info(f'Last Run: {datetime.fromtimestamp(info["lastRun"], get_localzone()).strftime("%m-%d-%Y at %H:%M:%S")}')
        else:
            lastRunColumn.info("Last Run: N/A")

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
                    try:
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
                    except Exception as e:
                        st.error(f"Failed to load thresholds file: {e}")
            else:
                st.warning(f"File not found: {file_path}")

    else:
        # Job not yet executed
        thresholdsFiles = []
        try:
            thresholdsFiles = [f[:-5] for f in os.listdir("../input/thresholds")]
        except Exception as e:
            st.error(f"Failed to list thresholds files: {e}")

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
                    try:
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
                    except Exception as e:
                        st.error(f"Failed to load thresholds file: {e}")
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