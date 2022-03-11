import os
from urllib import parse

import requests
import streamlit as st
from docker import APIClient

from FileHandler import openFile, openFolder
from utils.docker_utils import runConfigAssessmentTool, isDocker


def jobNotYetExecuted(client: APIClient, jobName: str, debug: bool, concurrentConnections: int):
    st.header(f"{jobName}")
    (
        openJobFileColumn,
        openThresholdsFileColumn,
        _,
        _,
        _,
    ) = st.columns([1.5, 2, 1.5, 3, 2])

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

    runColumn.text("")  # vertical padding
    if runColumn.button(f"Run", key=f"JobFile:{jobName}-Thresholds:{thresholds}-JobType:extract"):
        runConfigAssessmentTool(client, jobName, thresholds, debug, concurrentConnections)
