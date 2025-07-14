import json
import os
import re
import time

import streamlit as st
from docker import APIClient
from utils.streamlit_utils import rerun


def runConfigAssessmentTool(
    client: APIClient,
    jobFile: str,
    thresholds: str,
    debug: bool,
    concurrentConnections: int,
    username: str,
    password: str,
    auth_method: str,
    platformStr: str,
    tag: str,
):
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
        logTextBox.markdown(
            f"""
            <div style="height: 250px; overflow-y: auto; background-color: #f9f9f9; padding: 8px; border: 1px solid #ccc;" id="logbox">
                <pre style="margin: 0;">{logText}</pre>
            </div>
            <script>
                var logbox = document.getElementById('logbox');
                logbox.scrollTop = logbox.scrollHeight;
            </script>
            """,
            unsafe_allow_html=True
        )

    client.wait(container.get("Id"))

    # clear the live log box
    logTextBox.empty()

    # Display the final logs after container finishes
    st.info("Review the log above. When you're ready, click below to continue.")
    st.success("✅ Job finished successfully.")
    with st.expander("📄 View final logs", expanded=True):
        st.text_area("Log Output", logText, height=400)
        col1, col2 = st.columns([1, 2])
        with col1:
            if st.button("✅ Close Log Window"):
                st.experimental_rerun()
        with col2:
            st.caption("📋 To copy: Click inside the box, select All (Ctrl+A or ⌘+A),  then press Ctrl+C (or ⌘+C)")



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
