import json
import os
from datetime import datetime
from urllib import parse

import requests
import streamlit as st
from docker import APIClient
from FileHandler import openFile, openFolder
from tzlocal import get_localzone
from utils.docker_utils import isDocker, runConfigAssessmentTool


def jobPreviouslyExecuted(client: APIClient, jobName: str, debug: bool, concurrentConnections: int, platformStr: str, tag: str):
    st.header(f"{jobName}")
    info = json.loads(open(f"../output/{jobName}/info.json").read())

    (
        openOutputFolderColumn,
        openJobFileColumn,
        openThresholdsFileColumn,
        _,
    ) = st.columns([1.25, 1, 1.3, 1])

    openOutputFolderColumn.text("")  # vertical padding
    openOutputFolderColumn.text("")  # vertical padding
    if openOutputFolderColumn.button(f"Open Output Folder", key=f"{jobName}-outputFolder"):
        if not isDocker():
            openFolder(f"../output/{jobName}")
        else:
            payload = {"type": "folder", "path": f"output/{jobName}"}
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:16225?{payload}")

    openJobFileColumn.text("")  # vertical padding
    openJobFileColumn.text("")  # vertical padding
    if openJobFileColumn.button(f"Open JobFile", key=f"{jobName}-jobfile"):
        if not isDocker():
            openFile(f"../input/jobs/{jobName}.json")
        else:
            payload = {"type": "file", "path": f"input/jobs/{jobName}.json"}
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:16225?{payload}")

    thresholdsColumn, lastRunColumn, runColumn = st.columns([1, 1, 0.3])

    lastRunColumn.text("")  # vertical padding
    lastRunColumn.info(f'Last Run: {datetime.fromtimestamp(info["lastRun"], get_localzone()).strftime("%m-%d-%Y at %H:%M:%S")}')

    thresholdsFiles = [f[: len(f) - 5] for f in os.listdir("../input/thresholds")]
    if info["thresholds"] in thresholdsFiles:
        default_idx = thresholdsFiles.index(info["thresholds"])
    else:
        default_idx = 0
    thresholds = thresholdsColumn.selectbox(
        "Specify Thresholds File",
        thresholdsFiles,
        index=default_idx,
        key=f"{jobName}-new",
    )

    openThresholdsFileColumn.text("")  # vertical padding
    openThresholdsFileColumn.text("")  # vertical padding
    if openThresholdsFileColumn.button(f"Open Thresholds File", key=f"{jobName}-thresholds"):
        if not isDocker():
            openFile(f"../input/thresholds/{thresholds}.json")
        else:
            payload = {"type": "file", "path": f"input/thresholds/{thresholds}.json"}
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:16225?{payload}")

    dynamicCredentials = st.expander("Pass credentials dynamically (optional)")
    dynamicCredentials.write("Attention, if you use this option, it will dynamically change credentials for ALL controllers on the job file!")
    usrNameCol, pwdCol, dynChckCol = dynamicCredentials.columns(3)
    newUsrName = usrNameCol.text_input(label="New Username", value="Jeff", key=f"JobFile:{jobName}-usrCol")
    newPwd = pwdCol.text_input(label="New Password", value="hunter1", type="password", key=f"JobFile:{jobName}-pwdCol")
    dynChckCol.text("")
    dynChckCol.text("")
    dynamicCheck = dynChckCol.checkbox("Dynamic Credentials", key=f"JobFile:{jobName}-chckCol")

    runColumn.text("")  # vertical padding
    if runColumn.button(f"Run", key=f"JobFile:{jobName}-Thresholds:{thresholds}-JobType:extract"):
        username = newUsrName if dynamicCheck else None
        password = newPwd if dynamicCheck else None
        runConfigAssessmentTool(client, jobName, thresholds, debug, concurrentConnections, username, password, platformStr, tag)

    (
        openReportColumn,
        openReportButton,
    ) = st.columns([4, 1])

    reportFiles = [f[: len(f) - 5] for f in os.listdir(f"../output/{jobName}") if f.endswith("xlsx") and not f.startswith("~$")]
    report = openReportColumn.selectbox(
        "Specify Report to Open",
        reportFiles,
        key=f"{jobName}-open-report",
    )

    openReportButton.text("")  # vertical padding
    openReportButton.text("")  # vertical padding
    if openReportButton.button(f"Open Report", key=f"{jobName}-open-report-{report}"):
        if not isDocker():
            openFile(f"../output/{jobName}/{report}.xlsx")
        else:
            payload = {
                "type": "file",
                "path": f"output/{jobName}/{report}.xlsx",
            }
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:16225?{payload}")
