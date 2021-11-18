import os
from urllib import parse

import requests
import streamlit as st
from docker import APIClient

from FileHandler import openFile
from utils.docker_utils import runConfigAssessmentTool, isDocker


def jobNotYetExecuted(client: APIClient, jobName: str, debug: bool):
    st.header(f"{jobName}")
    c1, c2, c3 = st.columns(3)
    if c1.button(f"Open JobFile", key=f"{jobName}-jobfile"):
        if not isDocker():
            openFile(f"../input/jobs/{jobName}.json")
        else:
            payload = {"type": "file", "path": f"input/jobs/{jobName}.json"}
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:1337?{payload}")

    thresholdsColumn, warningColumn = st.columns(2)
    thresholdsFiles = [f[: len(f) - 5] for f in os.listdir("../input/thresholds")]
    thresholds = thresholdsColumn.selectbox("Specify Thresholds File", thresholdsFiles, key=f"{jobName}-new")

    if c3.button(f"Run", key=f"JobFile:{jobName}-Thresholds:{thresholds}-JobType:extract"):
        runConfigAssessmentTool(client, jobName, thresholds, debug)

    warningColumn.warning(f"Job has not yet been run")

    if c2.button(f"Open Thresholds File", key=f"{jobName}-thresholds"):
        if not isDocker():
            openFile(f"../input/thresholds/{thresholds}.json")
        else:
            payload = {"type": "file", "path": f"input/thresholds/{thresholds}.json"}
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:1337?{payload}")
