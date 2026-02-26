import os
from pathlib import Path
import streamlit as st

from backend.util.logging_utils import initLogging
from views.header import header
from views.jobHandler import jobHandler


def main():
    if not os.path.exists("output"):
        os.makedirs("output")

    debug, throttleNetworkConnections = header()

    # Initialize logging for the frontend
    initLogging(debug)

    if throttleNetworkConnections:
        concurrentNetworkConnections = st.sidebar.number_input("Concurrent Network Connections", min_value=1, max_value=100, value=50)
    else:
        concurrentNetworkConnections = 50

    orderedJobs = []
    if os.path.exists("input/jobs"):
        for jobName in os.listdir("input/jobs"):
            if jobName.startswith(".") or not jobName.endswith(".json"):
                continue

            jobName = jobName[: len(jobName) - 5]
            if Path(f"output/{jobName}/info.json").exists():
                orderedJobs.insert(0, jobName)
            else:
                orderedJobs.append(jobName)

    for jobName in orderedJobs:
        # Directly call jobHandler without Docker image tags
        jobHandler(jobName,debug, concurrentNetworkConnections)
        st.markdown("""---""")

main()