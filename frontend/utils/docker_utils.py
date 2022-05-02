import json
import re
import time

import streamlit as st
import os

from docker import APIClient

from utils.streamlit_utils import rerun


def runConfigAssessmentTool(client: APIClient, jobFile: str, thresholds: str, debug: bool, concurrentConnections: int, password_dynamically: bool, platformStr: str, tag: str):
    if not isDocker():
        root = os.path.abspath("..")
    else:
        root = os.environ["HOST_ROOT"]

    inputSource = f"{root}/input"
    outputSource = f"{root}/output"
    logsSource = f"{root}/logs"

    if os.name == "nt" or root[1] == ":":
        inputSource = ("/" + inputSource[:1] + "/" + inputSource[3:]).replace("\\", "/")
        outputSource = ("/" + outputSource[:1] + "/" + outputSource[3:]).replace("\\", "/")
        logsSource = ("/" + logsSource[:1] + "/" + logsSource[3:]).replace("\\", "/")

    command = ["-j", jobFile, "-t", thresholds, "-c", str(concurrentConnections), "-p", str(password_dynamically)]
    if debug:
        command.append("-d")
#-----------------Was used in previous idea-----------------------------------------
    # if password_dynamically:
    #     command.append("-p")

    container = client.create_container(
        image=f"ghcr.io/appdynamics/config-assessment-tool-backend-{platformStr}:{tag}",
        name=f"config-assessment-tool-backend-{platformStr}",
        volumes=[inputSource, outputSource],
        host_config=client.create_host_config(
            auto_remove=True,
            binds={
                inputSource: {
                    "bind": "/input",
                    "mode": "rw",
                },
                outputSource: {
                    "bind": "/output",
                    "mode": "rw",
                },
                logsSource: {
                    "bind": "/logs",
                    "mode": "rw",
                },
            },
        ),
        command=command,
    )
    client.start(container)

    logTextBox = st.empty()
    logText = ""
    for log in client.logs(container.get("Id"), stream=True):
        logText = log.decode("ISO-8859-1") + logText
        logTextBox.text_area("", logText, height=250)

    # small delay to see job ended
    time.sleep(5)
    # refresh the page to see newly generated report
    rerun()


def buildConfigAssessmentToolImage(client: APIClient, platformStr: str, tag: str):
    st.write(f"Docker image config_assessment_tool:latest not found. Please build the image.")
    if st.button(f"Build Image"):
        logTextBox = st.empty()
        logText = ""

        for output in client.build(
            path=os.path.abspath("../backend"),
            tag=f"appdynamics/config-assessment-tool-backend-{platformStr}:{tag}",
        ):
            output = output.decode("ISO-8859-1").strip("\r\n")
            for match in re.finditer(r"({.*})+", output):
                try:
                    logText = json.loads(match.group(1))["stream"] + logText
                except KeyError:
                    logText = match.group(1) + logText
                logTextBox.text_area("", logText, height=450)

        # small delay to see build ended
        time.sleep(5)
        # refresh the page
        rerun()


def isDocker():
    path = "/proc/self/cgroup"
    return os.path.exists("/.dockerenv") or os.path.isfile(path) and any("docker" in line for line in open(path))


def getImage(client: APIClient, imageName: str):
    return next(
        iter([image for image in client.images() if image["RepoTags"] is not None and any(tag for tag in image["RepoTags"] if tag == imageName)]),
        None,
    )
