#!/usr/bin/env python3
import logging
import os
import platform
import subprocess
import sys
import time
import zipfile
from http.client import RemoteDisconnected
from platform import uname
from urllib.error import URLError
from urllib.request import urlopen


def run(path: str, platformStr: str, tag: str):
    splash = """
                          __ _                                                              _        _              _
          ___ ___  _ __  / _(_) __ _        __ _ ___ ___  ___  ___ ___ _ __ ___   ___ _ __ | |_     | |_ ___   ___ | |
         / __/ _ \| '_ \| |_| |/ _` |_____ / _` / __/ __|/ _ \/ __/ __| '_ ` _ \ / _ \ '_ \| __|____| __/ _ \ / _ \| |
        | (_| (_) | | | |  _| | (_| |_____| (_| \__ \__ \  __/\__ \__ \ | | | | |  __/ | | | ||_____| || (_) | (_) | |
         \___\___/|_| |_|_| |_|\__, |      \__,_|___/___/\___||___/___/_| |_| |_|\___|_| |_|\__|     \__\___/ \___/|_|
                               |___/
        """
    logging.info(splash)

    # Check if config-assessment-tool images exist
    if (
        runBlockingCommand(f"docker images -q ghcr.io/appdynamics/config-assessment-tool-frontend-{platformStr}:{tag}") == ""
        or runBlockingCommand(f"docker images -q ghcr.io/appdynamics/config-assessment-tool-backend-{platformStr}:{tag}") == ""
    ):
        logging.info("Necessary Docker images not found.")
        build(platformStr, tag)
    else:
        logging.info("Necessary Docker images found.")

    # stop FileHandler
    logging.info("Terminating FileHandler if already running")
    if sys.platform == "win32":
        runBlockingCommand("WMIC path win32_process Where \"name like '%python%' and CommandLine like '%FileHandler.py%'\" CALL TERMINATE")
    else:
        runBlockingCommand("pgrep -f 'FileHandler.py' | xargs kill")

    # stop config-assessment-tool-frontend
    logging.info(f"Terminating config-assessment-tool-frontend-{platformStr} container if already running")
    containerId = runBlockingCommand("docker ps -f name=config-assessment-tool-frontend-" + platformStr + ' --format "{{.ID}}"')
    if containerId:
        runBlockingCommand(f"docker container stop {containerId}")

    # start FileHandler
    logging.info("Starting FileHandler")
    runNonBlockingCommand(f"{sys.executable}" + " frontend/FileHandler.py")

    # wait for file handler to start
    while True:
        logging.info("Waiting for FileHandler to start on http://localhost:16225")
        try:
            if urlopen("http://localhost:16225/ping").read() == b"pong":
                logging.info("FileHandler started")
                break
        except URLError:
            pass
        time.sleep(1)

    # start config-assessment-tool-frontend
    logging.info(f"Starting config-assessment-tool-frontend-{platformStr} container")
    runNonBlockingCommand(
        f"docker run "
        f'--name "config-assessment-tool-frontend-{platformStr}" '
        f"-v /var/run/docker.sock:/var/run/docker.sock "
        f'-v "{path}/logs:/logs" '
        f'-v "{path}/output:/output" '
        f'-v "{path}/input:/input" '
        f'-e HOST_ROOT="{path}" '
        f'-e PLATFORM_STR="{platformStr}" '
        f'-e TAG="{tag}" '
        f"-p 8501:8501 "
        f"--rm "
        f"ghcr.io/appdynamics/config-assessment-tool-frontend-{platformStr}:{tag} &"
    )

    # wait for config-assessment-tool-frontend to start
    while True:
        logging.info(f"Waiting for config-assessment-tool-frontend-{platformStr}:{tag} to start")
        try:
            if urlopen("http://localhost:8501").status == 200:
                logging.info(f"config-assessment-tool-frontend-{platformStr}:{tag} started")
                break
        except (URLError, RemoteDisconnected):
            pass
        time.sleep(1)

    # open web browser platform specific
    if platformStr == "windows":  # windows
        runBlockingCommand("start http://localhost:8501")
    elif "microsoft" in uname().release.lower():  # wsl
        runBlockingCommand("wslview http://localhost:8501")
    elif "mac" in platformStr:  # mac
        runBlockingCommand("open http://localhost:8501")
    elif platformStr == "linux":  # linux
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
        logging.info(f"Terminating config-assessment-tool-frontend-{platformStr}:{tag} container if still running")
        containerId = runBlockingCommand(f"docker ps -f name=config-assessment-tool-frontend-{platformStr}:{tag}" + ' --format "{{.ID}}"')
        if containerId:
            runBlockingCommand(f"docker container stop {containerId}")


# build docker images from source
def build(platform: str, tag: str):
    if os.path.isfile("backend/Dockerfile") and os.path.isfile("frontend/Dockerfile"):
        logging.info(f"Building ghcr.io/appdynamics/config-assessment-tool-backend-{platformStr}:{tag} from Dockerfile")
        runBlockingCommand(f"docker build --no-cache -t ghcr.io/appdynamics/config-assessment-tool-backend-{platformStr}:{tag} -f backend/Dockerfile .")
        logging.info(f"Building ghcr.io/appdynamics/config-assessment-tool-frontend-{platformStr}:{tag} from Dockerfile")
        runBlockingCommand(f"docker build --no-cache -t ghcr.io/appdynamics/config-assessment-tool-frontend-{platformStr}:{tag} -f frontend/Dockerfile .")
    else:
        logging.info("Dockerfiles not found in either backend/ or frontend/.")
        logging.info("Please either clone the full repository to build the images manually.")

    # Check if config-assessment-tool images exist
    if (
        runBlockingCommand(f"docker images -q ghcr.io/appdynamics/config-assessment-tool-frontend-{platformStr}:{tag}") == ""
        or runBlockingCommand(f"docker images -q ghcr.io/appdynamics/config-assessment-tool-backend-{platformStr}:{tag}") == ""
    ):
        logging.info("Failed to build Docker images.")
        sys.exit(1)


# pull latest images from ghrc.io if on a unix system
def pull(platformStr: str, tag: str):
    logging.info(f"Pulling ghcr.io/appdynamics/config-assessment-tool-backend-{platformStr}:latest")
    runBlockingCommand(f"docker pull ghcr.io/appdynamics/config-assessment-tool-backend-{platformStr}:latest")
    logging.info(f"Pulling ghcr.io/appdynamics/config-assessment-tool-frontend-{platformStr}:latest")
    runBlockingCommand(f"docker pull ghcr.io/appdynamics/config-assessment-tool-frontend-{platformStr}:latest")


# package minimal required files
def package():
    logging.info("Creating zip file")
    with zipfile.ZipFile("config-assessment-tool-dist.zip", "w") as zip_file:
        zip_file.write("README.md")
        zip_file.write("VERSION")
        zip_file.write("bin/config-assessment-tool.py")
        zip_file.write("input/jobs/DefaultJob.json")
        zip_file.write("input/thresholds/DefaultThresholds.json")
        zip_file.write("frontend/FileHandler.py")
    logging.info("Created config-assessment-tool-dist.zip")


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
def verifySoftwareVersion() -> str:
    if sys.platform == "win32":
        latestTag = runBlockingCommand(
            'powershell -Command "(Invoke-WebRequest https://api.github.com/repos/appdynamics/config-assessment-tool/tags | ConvertFrom-Json)[0].name"'
        )
    else:
        latestTag = runBlockingCommand(
            "curl -s https://api.github.com/repos/appdynamics/config-assessment-tool/tags | grep 'name' | head -n 1 | cut -d ':' -f 2 | cut -d '\"' -f 2"
        )

    logging.info(f"Latest release tag from https://api.github.com/repos/appdynamics/config-assessment-tool/tags is {latestTag}")

    # get local tag from VERSION file
    localTag = "unknown"
    if os.path.isfile("VERSION"):
        with open("VERSION", "r") as versionFile:
            localTag = versionFile.read().strip()

    if latestTag != localTag:
        logging.warning(f"You are using an outdated version of the software. Current {localTag} Target {latestTag}")
        logging.warning("You can get the latest version from https://github.com/Appdynamics/config-assessment-tool/releases")
    else:
        logging.info(f"You are using the latest version of the software. Current {localTag}")

    return localTag


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

    logging.info(f"Working directory is {os.getcwd()}")

    tag = verifySoftwareVersion()

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
        logging.error(f"Unsupported platform {sys.platform}")
        platformStr = "unknown"
        exit(1)

    logging.info(f"Platform: {platformStr}")

    # parse command line arguments
    if len(sys.argv) == 1 or sys.argv[1] == "--help":
        msg = """
    Usage: config-assessment-tool.py [OPTIONS]
    Options:
      --run, Run the config-assessment-tool
      --build, Build frontend and backend from Dockerfile
      --package, Create lightweight package for distribution
      --help, Show this message and exit.
              """.strip()
        print(msg)
        sys.exit(1)
    if sys.argv[1] == "--run":
        run(path, platformStr, tag)
    elif sys.argv[1] == "--build":
        build(platformStr, tag)
    elif sys.argv[1] == "--package":
        package()
    else:
        print(f"Unknown option: {sys.argv[1]}")
        print("Use --help for usage information")
        sys.exit(1)

    sys.exit(0)
