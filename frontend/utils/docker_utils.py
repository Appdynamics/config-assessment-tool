import json
import os
import platform
import re
import time

import streamlit as st
from docker import APIClient


def runConfigAssessmentTool(client: APIClient, jobFile: str, thresholds: str,
                            debug: bool, concurrentConnections: int,
                            username: str, password: str, auth_method: str):
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

    command = ["-j", jobFile, "-t", thresholds, "-c", str(concurrentConnections)]
    if debug:
        command.append("-d")
    if username:
        command.extend(["-u", username])
    if password:
        command.extend(["-p", password])
    if auth_method:
        command.extend(["-m", auth_method])


    arch = get_arch()
    backend_image_tag, _ = get_image_tags()
    container = client.create_container(
        image=backend_image_tag,
        name=f"config-assessment-tool-backend-{arch}",
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
    time.sleep(8)
    # refresh the page to see newly generated report
    from utils.streamlit_utils import rerun
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


def get_image_tags():
    arch = get_arch()
    version = get_version()
    backend_tag = f"ghcr.io/appdynamics/config-assessment-tool-backend-{arch}:{version}"
    frontend_tag = f"ghcr.io/appdynamics/config-assessment-tool-frontend-{arch}:{version}"
    return backend_tag, frontend_tag


def get_arch():
    uname_s = platform.system()
    uname_m = platform.machine()
    os_part = "unknown_os"
    arch_part = "unknown_arch"

    if uname_s == "Darwin":
        os_part = "macos"
    elif uname_s == "Linux":
        os_part = "linux"
    elif "CYGWIN" in uname_s or "MINGW" in uname_s or "MSYS" in uname_s:
        os_part = "windows"

    if uname_m in ["x86_64"]:
        arch_part = "x86"
    elif uname_m in ["arm64", "aarch64"]:
        arch_part = "arm"

    return f"{os_part}-{arch_part}"


def get_version():
    # Find the project base directory (two levels up from this file)
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
    version_path = os.path.join(base_dir, 'VERSION')
    try:
        with open(version_path) as f:
            return f.read().strip()
    except Exception:
        return "unknown"
