#!/usr/bin/env python3
import logging
import os
import subprocess
import sys
import time
import zipfile
from platform import uname

import requests


def run(path: str):
    splash = """
                          __ _                                                              _        _              _
          ___ ___  _ __  / _(_) __ _        __ _ ___ ___  ___  ___ ___ _ __ ___   ___ _ __ | |_     | |_ ___   ___ | |
         / __/ _ \| '_ \| |_| |/ _` |_____ / _` / __/ __|/ _ \/ __/ __| '_ ` _ \ / _ \ '_ \| __|____| __/ _ \ / _ \| |
        | (_| (_) | | | |  _| | (_| |_____| (_| \__ \__ \  __/\__ \__ \ | | | | |  __/ | | | ||_____| || (_) | (_) | |
         \___\___/|_| |_|_| |_|\__, |      \__,_|___/___/\___||___/___/_| |_| |_|\___|_| |_|\__|     \__\___/ \___/|_|
                               |___/
        """
    logging.info(splash)

    # stop FileHandler
    logging.info("Terminating FileHandler if already running")
    if sys.platform == "win32":
        runBlockingCommand("WMIC path win32_process Where \"name like '%python%' and CommandLine like '%FileHandler.py%'\" CALL TERMINATE")
    else:
        runBlockingCommand("pgrep -f 'FileHandler.py' | xargs kill")

    # stop config-assessment-tool-frontend
    logging.info("Terminating config-assessment-tool-frontend container if already running")
    containerId = runBlockingCommand('docker ps -f name=config-assessment-tool-frontend --format "{{.ID}}"')
    if containerId:
        runBlockingCommand(f"docker container stop {containerId}")

    # start FileHandler
    logging.info("Starting FileHandler")
    runNonBlockingCommand("python frontend/FileHandler.py")

    # wait for file handler to start
    while True:
        logging.info("Waiting for FileHandler to start")
        if requests.get("http://localhost:1337/ping").text == "pong":
            break
        time.sleep(1)

    # start config-assessment-tool-frontend
    logging.info("Starting config-assessment-tool-frontend container")
    runNonBlockingCommand(
        f"docker run "
        f'--name "config-assessment-tool-frontend" '
        f"-v /var/run/docker.sock:/var/run/docker.sock "
        f"-v {path}/logs:/logs "
        f"-v {path}/output:/output "
        f"-v {path}/input:/input "
        f'-e HOST_ROOT="{path}" '
        f"-p 8501:8501 "
        f"--rm "
        f"ghcr.io/appdynamics/config-assessment-tool-frontend:latest &"
    )

    # wait for config-assessment-tool-frontend to start
    while True:
        logging.info("Waiting for config-assessment-tool-frontend to start")
        if "config-assessment-tool-frontend" in runBlockingCommand("docker ps -a --format '{{.Names}}'"):
            logging.info("config-assessment-tool-frontend started")
            break
        time.sleep(1)

    # open web browser platform specific
    if sys.platform == "win32":  # windows
        runBlockingCommand("start http://localhost:8501")
    elif "microsoft" in uname().release.lower():  # wsl
        runBlockingCommand("wslview http://localhost:8501")
    elif sys.platform == "darwin":  # mac
        runBlockingCommand("open http://localhost:8501")
    elif sys.platform == "linux":  # linux
        runBlockingCommand("xdg-open http://localhost:8501")
    else:
        logging.info("Unsupported platform, trying to open web browser to http://localhost:8501")
        runBlockingCommand("open http://localhost:8501")

    # Loop until user exits
    logging.info("Press Ctrl-C to stop the config-assessment-tool")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        logging.info("Received KeyboardInterrupt")
        logging.info("Terminating config-assessment-tool-frontend container if still running")
        containerId = runBlockingCommand('docker ps -f name=config-assessment-tool-frontend --format "{{.ID}}"')
        if containerId:
            runBlockingCommand(f"docker container stop {containerId}")


def build():
    if os.path.isfile("backend/Dockerfile") and os.path.isfile("frontend/Dockerfile"):
        logging.info("Building ghcr.io/appdynamics/config-assessment-tool-backend:latest from Dockerfile")
        runBlockingCommand("docker build -t ghcr.io/appdynamics/config-assessment-tool-backend:latest -f backend/Dockerfile .")
        logging.info("Building ghcr.io/appdynamics/config-assessment-tool-frontend:latest from Dockerfile")
        runBlockingCommand("docker build -t ghcr.io/appdynamics/config-assessment-tool-frontend:latest -f frontend/Dockerfile .")
    else:
        logging.info("Dockerfiles not found in either backend/ or frontend/.")


def pull():
    if sys.platform == "win32":
        logging.info("Currently only pre-built images are available for *nix systems.")
        logging.info("Please build the images from source with the --build command.")
    else:
        logging.info("Pulling ghcr.io/appdynamics/config-assessment-tool-backend:latest")
        runBlockingCommand("docker pull ghcr.io/appdynamics/config-assessment-tool-backend:latest")
        logging.info("Pulling ghcr.io/appdynamics/config-assessment-tool-frontend:latest")
        runBlockingCommand("docker pull ghcr.io/appdynamics/config-assessment-tool-frontend:latest")


def package():
    logging.info("Creating zip file")
    with zipfile.ZipFile("config-assessment-tool.zip", "w") as zip_file:
        zip_file.write("README.md")
        zip_file.write("bin/config-assessment-tool.py")
        zip_file.write("input/jobs/DefaultJob.json")
        zip_file.write("input/thresholds/DefaultThresholds.json")
        zip_file.write("frontend/FileHandler.py")
    logging.info("Created config-assessment-tool.zip")


def runBlockingCommand(command: str):
    output = ""
    with subprocess.Popen(command, stdout=subprocess.PIPE, stderr=None, shell=True) as process:
        line = process.communicate()[0].decode("utf-8").strip()
        if line:
            logging.debug(line)
        output += line
    return output.strip()


def runNonBlockingCommand(command: str):
    subprocess.Popen(command, stdout=None, stderr=None, shell=True)


if __name__ == "__main__":
    assert sys.version_info >= (3, 5), "Python 3.5 or higher required"

    path = os.path.realpath(f"{__file__}/../..")
    os.chdir(path)

    # create logs and output directories
    if not os.path.exists("logs"):
        os.makedirs("logs")
    if not os.path.exists("output"):
        os.makedirs("output")

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler("logs/config-assessment-tool-frontend.log"),
            logging.StreamHandler(),
        ],
    )

    if sys.argv[1] == "--run":
        run(path)
    elif sys.argv[1] == "--build":
        build()
    elif sys.argv[1] == "--pull":
        pull()
    elif sys.argv[1] == "--package":
        package()
    elif len(sys.argv) == 1 or sys.argv[1] == "--help":
        msg = """
    Usage: config-assessment-tool.py [OPTIONS]
    Options:
      --run, Run the config-assessment-tool
      --build, Build frontend and backend from Dockerfile
      --pull, Pull latest images from GitHub
      --package, Create lightweight package for distribution
      --help, Show this message and exit.
              """.strip()
        logging.info(msg)
        sys.exit(1)

    else:
        logging.error(f"Unknown option: {sys.argv[1]}")
        logging.info("Use --help for usage information")
        sys.exit(1)

    sys.exit(0)
