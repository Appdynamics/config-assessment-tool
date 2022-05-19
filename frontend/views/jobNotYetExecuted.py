import os
from urllib import parse

import requests
import streamlit as st
from docker import APIClient
from FileHandler import openFile, openFolder
from utils.docker_utils import isDocker, runConfigAssessmentTool


def jobNotYetExecuted(
    client: APIClient, jobName: str, debug: bool, concurrentConnections: int, username: str, password: bool, platformStr: str, tag: str
):
    st.header(f"{jobName}")
    (
        openJobFileColumn,
        openThresholdsFileColumn,
        _,
    ) = st.columns([1, 2, 2])

    openJobFileColumn.text("")  # vertical padding
    openJobFileColumn.text("")  # vertical padding
    if openJobFileColumn.button(f"Open JobFile", key=f"{jobName}-jobfile"):
        if not isDocker():
            openFile(f"../input/jobs/{jobName}.json")
        else:
            payload = {"type": "file", "path": f"input/jobs/{jobName}.json"}
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:16225?{payload}")

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
        if not isDocker():
            openFile(f"../input/thresholds/{thresholds}.json")
        else:
            payload = {"type": "file", "path": f"input/thresholds/{thresholds}.json"}
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:16225?{payload}")

    dynamicCredentials = st.expander("Use different credentials (This is optional!)")
    dynamicCredentials.write("Attention, if you use this option, it will dynamicly change credentials for ALL controllers on the job file!")
    usrNameCol, pwdCol, dynChckCol = dynamicCredentials.columns(3)
    newUsrName = usrNameCol.text_input(label="New Username", value="Jeff", key=f"JobFile:{jobName}-username")
    newPwd = pwdCol.text_input(label="New Password", value="examplepwd", type="password", key=f"JobFile:{jobName}-pwd")
    dynChckCol.text("")
    dynChckCol.text("")
    dynamicCheck = dynChckCol.checkbox("Dynamic Credentials", key=f"JobFile:{jobName}-checkbox")

    runColumn.text("")  # vertical padding
    if runColumn.button(f"Run", key=f"JobFile:{jobName}-Thresholds:{thresholds}-JobType:extract"):
        username = newUsrName if dynamicCheck else None
        password = newPwd if dynamicCheck else None  # I changed from "None" to None, but not sure if it's a correct
        # implemenation of this tho

        runConfigAssessmentTool(client, jobName, thresholds, debug, concurrentConnections, username, password, platformStr, tag)
