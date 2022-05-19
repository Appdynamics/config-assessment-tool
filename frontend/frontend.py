import os
import platform
import sys
from pathlib import Path
from platform import uname

import docker
import requests
import streamlit as st
from utils.docker_utils import getImage, isDocker
from views.header import header
from views.jobNotYetExecuted import jobNotYetExecuted
from views.jobPreviouslyExecuted import jobPreviouslyExecuted


def main():
    client = docker.from_env().api

    if not os.path.exists("../output"):
        os.makedirs("../output")

    # create page header, header contains debug checkbox
    debug, throttleNetworkConnections = header()

    if throttleNetworkConnections:
        concurrentNetworkConnections = st.sidebar.number_input("Concurrent Network Connections", min_value=1, max_value=100, value=50)
    else:
        concurrentNetworkConnections = 50

    """-----The dynamic stuff-----"""
    username = None
    password = None

    if not isDocker():
        # determine platform
        if sys.platform.lower() == "win32":  # windows
            platformStr = "windows"
        elif "microsoft" in uname().release.lower():  # wsl
            platformStr = "linux"
        elif sys.platform.lower() == "darwin":  # mac
            platformStr = "mac"
            if platform.processor() == "arm":
                platformStr = "mac-m1"
        elif sys.platform.lower() == "linux":  # linux
            platformStr = "linux"
        else:
            st.write.error(f"Unsupported platform {sys.platform}")
            platformStr = "unknown"
            exit(1)

        # determine tag
        # get local tag from VERSION file
        tag = "unknown"
        if os.path.isfile("../VERSION"):
            with open("../VERSION", "r") as versionFile:
                tag = versionFile.read().strip()
    else:
        platformStr = os.environ["PLATFORM_STR"]
        tag = os.environ["TAG"]

    # does docker image 'config_assessment_tool:latest' exist
    if getImage(client, f"ghcr.io/appdynamics/config-assessment-tool-backend-{platformStr}:{tag}") is None:
        st.write(f"Image config-assessment-tool-backend-{platformStr}:{tag} not found")
        st.write(f"Please either build from source with --build")
        st.write(f"In order to --build you will need to download the full source")
    else:
        # order jobs which have already been ran at the top
        orderedJobs = []
        for jobName in os.listdir("../input/jobs"):
            if jobName.startswith("."):  # skip hidden files (.DS_Store)
                continue

            jobName = jobName[: len(jobName) - 5]  # strip .json
            if Path(f"../output/{jobName}/info.json").exists():
                orderedJobs.insert(0, jobName)
            else:
                orderedJobs.append(jobName)

        for jobName in orderedJobs:
            if Path(f"../output/{jobName}/info.json").exists():
                jobPreviouslyExecuted(client, jobName, debug, concurrentNetworkConnections, username, password, platformStr, tag)
            else:
                jobNotYetExecuted(client, jobName, debug, concurrentNetworkConnections, username, password, platformStr, tag)
            st.markdown("""---""")


main()
