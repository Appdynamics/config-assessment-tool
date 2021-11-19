import os

import docker
from pathlib import Path

import streamlit as st

from views.jobPreviouslyExecuted import jobPreviouslyExecuted
from views.jobNotYetExecuted import jobNotYetExecuted
from views.header import header


def main():
    client = docker.from_env().api

    if not os.path.exists("../output"):
        os.makedirs("../output")

    # create page header, header contains debug checkbox
    debug = header()

    # does docker image 'config_assessment_tool:latest' exist
    if (
        next(
            iter(
                [
                    image
                    for image in client.images()
                    if image["RepoTags"] is not None
                    and any(tag for tag in image["RepoTags"] if tag == "ghcr.io/appdynamics/config-assessment-tool-backend:latest")
                ]
            ),
            None,
        )
        is None
    ):
        st.write(f"Image config-assessment-tool-backend:latest not found")
        st.write(f"Please verify images were created successfully with bin/run.sh")
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
                jobPreviouslyExecuted(client, jobName, debug)
            else:
                jobNotYetExecuted(client, jobName, debug)
            st.markdown("""---""")


main()
