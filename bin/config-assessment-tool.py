#!/usr/bin/env python3
import logging
import os
import subprocess
import sys
import time
import zipfile
from http.client import RemoteDisconnected
from platform import uname
from urllib.error import URLError

from urllib.request import urlopen


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

    # Check if config-assessment-tool-frontend exists
    if (
        runBlockingCommand("docker images -q ghcr.io/appdynamics/config-assessment-tool-frontend:latest") == ""
        or runBlockingCommand("docker images -q ghcr.io/appdynamics/config-assessment-tool-backend:latest") == ""
    ):
        logging.error("Necessary Docker images not found.")
        logging.error("Please re-run with --pull or --build.")
        sys.exit(1)

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
        try:
            if urlopen("http://localhost:16225/ping").read() == b"pong":
                logging.info("FileHandler started")
                break
        except URLError:
            pass
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
        try:
            if urlopen("http://localhost:8501").status == 200:
                logging.info("config-assessment-tool-frontend started")
                break
        except (URLError, RemoteDisconnected):
            pass
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


# build docker images from source
def build():
    if os.path.isfile("backend/Dockerfile") and os.path.isfile("frontend/Dockerfile"):
        logging.info("Building ghcr.io/appdynamics/config-assessment-tool-backend:latest from Dockerfile")
        runBlockingCommand("docker build -t ghcr.io/appdynamics/config-assessment-tool-backend:latest -f backend/Dockerfile .")
        logging.info("Building ghcr.io/appdynamics/config-assessment-tool-frontend:latest from Dockerfile")
        runBlockingCommand("docker build -t ghcr.io/appdynamics/config-assessment-tool-frontend:latest -f frontend/Dockerfile .")
    else:
        logging.info("Dockerfiles not found in either backend/ or frontend/.")
        logging.info("Please either clone the full repository to build the images manually.")


# pull latest images from ghrc.io if on a unix system
def pull():
    if sys.platform == "win32":
        logging.info("Currently only pre-built images are available for *nix systems.")
        logging.info("Please build the images from source with the --build command.")
    else:
        logging.info("Pulling ghcr.io/appdynamics/config-assessment-tool-backend:latest")
        runBlockingCommand("docker pull ghcr.io/appdynamics/config-assessment-tool-backend:latest")
        logging.info("Pulling ghcr.io/appdynamics/config-assessment-tool-frontend:latest")
        runBlockingCommand("docker pull ghcr.io/appdynamics/config-assessment-tool-frontend:latest")


# package minimal required files
def package():
    logging.info("Creating zip file")
    with zipfile.ZipFile("config-assessment-tool.zip", "w") as zip_file:
        zip_file.write("README.md")
        zip_file.write("VERSION")
        zip_file.write("bin/config-assessment-tool.py")
        zip_file.write("input/jobs/DefaultJob.json")
        zip_file.write("input/thresholds/DefaultThresholds.json")
        zip_file.write("frontend/FileHandler.py")
    logging.info("Created config-assessment-tool.zip")


# execute blocking command, gather output and return it
def runBlockingCommand(command: str):
    output = ""
    with subprocess.Popen(command, stdout=subprocess.PIPE, stderr=None, shell=True) as process:
        line = process.communicate()[0].decode("ISO-8859-1").strip()
        if line:
            logging.debug(line)
        output += line
    return output.strip()


# execute non-blocking command, ignore output
def runNonBlockingCommand(command: str):
    subprocess.Popen(command, stdout=None, stderr=None, shell=True)


# verify current software version against GitHub release tags
def verifySoftwareVersion():
    if sys.platform == "win32":
        latestTag = runBlockingCommand(
            'powershell -Command "(Invoke-WebRequest https://api.github.com/repos/appdynamics/config-assessment-tool/tags | ConvertFrom-Json)[0].name"'
        )
    else:
        latestTag = runBlockingCommand(
            "curl -s https://api.github.com/repos/appdynamics/config-assessment-tool/tags | grep 'name' | head -n 1 | cut -d ':' -f 2 | cut -d '\"' -f 2"
        )

    # get local tag from VERSION file
    localTag = "unknown"
    if os.path.isfile("VERSION"):
        with open("VERSION", "r") as versionFile:
            localTag = versionFile.read().strip()

    if latestTag != localTag:
        logging.warning(f"You are using an outdated version of the software. Current {localTag} Target {latestTag}")
        logging.warning("You can get the latest version from https://github.com/Appdynamics/config-assessment-tool/releases")
    else:
        logging.info(f"Already up to date!")


if __name__ == "__main__":
    assert sys.version_info >= (3, 5), "Python 3.5 or higher required"

    # cd to config-assessment-tool root directory
    path = os.path.realpath(f"{__file__}/../..")
    os.chdir(path)

    # create logs and output directories
    if not os.path.exists("logs"):
        os.makedirs("logs")
    if not os.path.exists("output"):
        os.makedirs("output")

    # init logging
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler("logs/config-assessment-tool-frontend.log"),
            logging.StreamHandler(),
        ],
    )

    verifySoftwareVersion()

    # parse command line arguments
    if len(sys.argv) == 1 or sys.argv[1] == "--help":
        msg = """
    Usage: config-assessment-tool.py [OPTIONS]
    Options:
      --run, Run the config-assessment-tool
      --build, Build frontend and backend from Dockerfile
      --pull, Pull latest images from GitHub
      --package, Create lightweight package for distribution
      --help, Show this message and exit.
              """.strip()
        print(msg)
        sys.exit(1)
    if sys.argv[1] == "--run":
        run(path)
    elif sys.argv[1] == "--build":
        build()
    elif sys.argv[1] == "--pull":
        pull()
    elif sys.argv[1] == "--package":
        package()
    else:
        print(f"Unknown option: {sys.argv[1]}")
        print("Use --help for usage information")
        sys.exit(1)

    sys.exit(0)
