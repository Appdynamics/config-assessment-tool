import json
import os
import subprocess
from datetime import datetime
from urllib import parse

import requests
import streamlit as st
from docker import APIClient
from tzlocal import get_localzone

from FileHandler import openFile
from utils.docker_utils import runConfigAssessmentTool, isDocker


def jobPreviouslyExecuted(client: APIClient, jobName: str, debug: bool):
    st.header(f"{jobName}")
    info = json.loads(open(f"../output/{jobName}/info.json").read())

    (
        openJobFileColumn,
        openThresholdsFileColumn,
        openExcelReportColumn,
        runColumn,
    ) = st.columns(4)
    if openJobFileColumn.button(f"Open JobFile", key=f"{jobName}-jobfile"):
        if not isDocker():
            openFile(f"../input/jobs/{jobName}.json")
        else:
            payload = {"type": "file", "path": f"input/jobs/{jobName}.json"}
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:1337?{payload}")

    if openExcelReportColumn.button(f"Open Excel Report", key=f"{jobName}-excel"):
        if not isDocker():
            openFile(f"../output/{jobName}/{jobName}-Report.xlsx")
        else:
            payload = {
                "type": "file",
                "path": f"output/{jobName}/{jobName}-Report.xlsx",
            }
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:1337?{payload}")

    thresholdsColumn, lastRunColumn = st.columns(2)

    lastRunColumn.info(f'Last Run: {datetime.fromtimestamp(info["lastRun"], get_localzone()).strftime("%m-%d-%Y at %H:%M:%S")}')

    thresholdsFiles = [f[: len(f) - 5] for f in os.listdir("../input/thresholds")]
    default_idx = thresholdsFiles.index(info["thresholds"])
    thresholds = thresholdsColumn.selectbox(
        "Specify Thresholds File",
        thresholdsFiles,
        index=default_idx,
        key=f"{jobName}-new",
    )

    if openThresholdsFileColumn.button(f"Open Thresholds File", key=f"{jobName}-thresholds"):
        if not isDocker():
            openFile(f"input/thresholds/{thresholds}.json")
        else:
            payload = {"type": "file", "path": f"input/thresholds/{thresholds}.json"}
            payload = parse.urlencode(payload)
            requests.get(f"http://host.docker.internal:1337?{payload}")

    if runColumn.button(f"Run", key=f"JobFile:{jobName}-Thresholds:{thresholds}-JobType:extract"):
        runConfigAssessmentTool(client, jobName, thresholds, debug)
