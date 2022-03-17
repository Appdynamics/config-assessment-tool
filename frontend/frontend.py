import os

import docker
from pathlib import Path

import requests
import streamlit as st

from utils.docker_utils import getImage
from views.jobPreviouslyExecuted import jobPreviouslyExecuted
from views.jobNotYetExecuted import jobNotYetExecuted
from views.header import header


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

    # does docker image 'config_assessment_tool:latest' exist
    if getImage(client, "ghcr.io/appdynamics/config-assessment-tool-backend:latest") is None:
        st.write(f"Image config-assessment-tool-backend:latest not found")
        st.write(f"Please either build from source with --build or pull from ghrc with --pull")
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
                jobPreviouslyExecuted(client, jobName, debug, concurrentNetworkConnections)
            else:
                jobNotYetExecuted(client, jobName, debug, concurrentNetworkConnections)
            st.markdown("""---""")


main()
