# Config Assessment Tool - Beta Release Runbook (1.8.0-beta.1)

This document provides a quick-start guide for the 1.8.0-beta.1 release. This beta release introduces usage improvements, UI improvements, feature and bug fixes. 

**[View Full Release Notes](https://github.com/alexafshar/config-assessment-tool/releases/tag/1.8.0-beta.1)**

## Table of Contents
There are multiple ways to run the tool
- [1. Using Docker](#1-using-docker)
    - [Option A: Using Shell Script (Easier)](#option-a-using-shell-script-easier)
    - [Option B: Manual Docker Command (Advanced)](#option-b-manual-docker-command-advanced)
        - [Run the UI](#run-the-ui-frontend--backend)
        - [Run Headless](#run-headless-backend-only)
        - [Getting Help](#getting-help)
- [2. Platform Executable (Standalone Binary)](#2-platform-executable-standalone-binary)
    - [UI Mode](#ui-mode)
    - [Headless Mode](#headless-mode)
    - [Getting Help](#getting-help-1)
- [3. Running from source (`config-assessment-tool.sh`)](#3-running-from-source-config-assessment-toolsh)
    - [Prerequisites](#prerequisites)
    - [Usage](#usage)
    - [Plugin Management](#plugin-management)
    - [Shutdown](#shutdown)
    - [Getting Help](#getting-help-2)
- [4. Configuration (Important)](#4-configuration-important)
    - [Authentication Options (`authType`)](#authentication-options-authtype)

# 1. Using Docker
*Node: Docker instructions are for MacOS/Linux. Windows PowerShell runbook instructions is coming soon. Use executable bundle for Windows in the meantime or run from source*  

This is the easiest way to run the tool using Docker. The only prerequisite is a local installation of the Docker engine or Docker desktop and access to the shell Terminal.

## Option A: Using Shell Script (Easier)

1.  **Download or Clone the Source:**
    *   Download the Source Code (zip/tar.gz) from the **[Releases Page](https://github.com/alexafshar/config-assessment-tool/releases/tag/1.8.0-beta.1)**.
    *   *OR* if you have git installed, clone the repository: `git clone https://github.com/alexafshar/config-assessment-tool.git`

2.  **Unzip the source if downloaded. In a shell Terminal, navigate to the directory:**
    ```bash
    cd config-assessment-tool
    ```

**Run the Tool using the provided shell script in the downloaded/cloned source:**

**UI mode: Start the UI:**
```bash
./config-assessment-tool.sh docker --ui
```
**Headless mode: Run Headless:**
```bash
./config-assessment-tool.sh docker -j DefaultJob
```

## Option B: Manual Docker Command (Advanced)
If you prefer to run `docker` directly without downloading the source and running the provided shell script, you must provide the required docker command line options to map the local directories to the docker container directories so that CAT running in the container can find the job files or output and log files.

**Suggested Directory Structure:**
```text
configuration-assessment-tool/
├── input
│   ├── jobs
│   │   └── DefaultJob.json
│   └── thresholds
│       └── DefaultThresholds.json
├── logs
└── output
```
*You may download to use in your setup the following samples: [DefaultJob.json](input/jobs/DefaultJob.json) and [DefaultThresholds.json](input/thresholds/DefaultThresholds.json)*

### Run the UI (Frontend + Backend)
Assuming the above directory structure:
   ```bash
   cd config-assessment-tool
   ```
```bash
docker run \
  --name "cat-tool-beta" \
  --rm \
  -p 8501:8501 \
  -v ./input:/app/input \
  -v ./output:/app/output \
  -v ./logs:/app/logs \
  ghcr.io/alexafshar/config-assessment-tool:1.8.0-beta.1 --ui
```
*Access the UI at http://localhost:8501*

### Run Headless (Backend Only)
Pass backend arguments directly (e.g., `-j`).
```bash
docker run \
  --name "cat-tool-beta-backend" \
  --rm \
  -v "$(pwd)/input":/app/input \
  -v "$(pwd)/output":/app/output \
  -v "$(pwd)/logs":/app/logs \
  ghcr.io/alexafshar/config-assessment-tool:1.8.0-beta.1 \
  -j <job-file-name>
```

### Getting Help
Display the help menu for the container.
```bash
docker run ghcr.io/alexafshar/config-assessment-tool:1.8.0-beta.1 --help
```

---

# 2. Platform Executable (Standalone Binary)

Download the executable bundle for your platform from the [Release Page](../../releases/tag/1.8.0-beta.1)

## UI Mode
*Note: Beta binaries contain the Web UI. You can launch it using `--ui` command line option 
1.  Run the executable with the `--ui` flag:
    *   **macOS**:
        > **Note:** macOS binaries need to go through Apple's certificaiton process before they run unhindered on your local macOS machine. To bypass certification you must remove the quarantine attribute from the downloaded bundle otherwise macOS will prevent it from running. This is a one time process when you download a new bundle.
        1. Unzip the downloaded bundle and open a Terminal and navigate to the directory containing the extracted files.
        2. Navigate into the directory if you haven't already:
            ```bash
            cd config-assessment-tool-macosx-*
            ```
        2.  Run the following command to allow the application to run:
            ```bash
            sudo xattr -rd com.apple.quarantine .     # this removes quarantine attribute for all files in current directory
            ```
2.  Navigate into the directory of extracted bundle if you haven't already:
    ```bash
    cd config-assessment-tool-<platform>-1.8.0-beta.1  # platform is specific to your OS
    ```
3.  Look for the executable file ```config-assessment-tool``` and run the tool with the UI flag:
    ```bash
      ./config-assessment-tool --ui
    ```
    For example:
       *   **macOS**: `./config-assessment-tool --ui`
       *   **Linux**: `./config-assessment-tool --ui`
       *   **Windows**: `PS> .\config-assessment-tool.exe --ui`
<br><br> 
4. Open your browser to the URL displayed in the console (typically `http://localhost:8501`). If running on a remote server, replace `localhost` with the server's IP address.

## Headless Mode
1.  Navigate to the directory where you extracted the release.
2.  Run the executable:
    *   **macOS**:
        > **Note:** macOS binaries need to go through Apple's certificaiton process before they run unhindered on your local macOS machine. To bypass certification you must remove the quarantine attribute from the downloaded bundle otherwise macOS will prevent it from running. This is a one time process when you download a new bundle. 
        1.  Open a Terminal and navigate to the directory containing the unzipped folder.
        2.  Run the following command to allow the application to run:
            ```bash
            sudo xattr -rd com.apple.quarantine .     # this removes quarantine attribute for all files in current directory
            ```
        3.  Navigate into the directory:
            ```bash
            cd config-assessment-tool-macosx-*
            ```
        4.  Run the tool:
            ```bash
            ./config-assessment-tool -j DefaultJob
            ```
    For example:
    *   **macOS**: `./config-assessment-tool -j <job-file>`
    *   **Linux**: `./config-assessment-tool -j <job-file>`
    *   **Windows**: `PS> .\config-assessment-tool.exe -j <job-file>`

## Getting Help
Display the available command line arguments.
*   **macOS**: `./config-assessment-tool --help`
*   **Linux**: `./config-assessment-tool --help`
*   **Windows**: `.\config-assessment-tool.exe --help`

---

# 3. Running from source (`config-assessment-tool.sh`) 

The `config-assessment-tool.sh` script is the recommended way to run the tool on macOS and Linux once you have the source code download or cloned locally. 

## Prerequisites
*   **Docker** If using the script to run the docker image (recommended for most users)
*   *OR* Python 3.12+ with Pipenv if running from source without docker (advanced users or those with specific custom needs)

## Usage
**Get the source code** 
- Download the Source Code (zip/tar.gz) from the **[Releases Page](https://github.com/alexafshar/config-assessment-tool/releases/tag/1.8.0-beta.1)**.
- *OR* if you have git installed, clone the repository: `git clone https://github.com/alexafshar/config-assessment-tool.git`

**Unzip the source if necessary. In a shell Terminal, navigate to the directory of expanded source**

```bash
cd config-assessment-tool
```

**Start the User Interface (UI):**
   - This will launch the Web UI in your default browser where you can select which job to run.
```bash
# Using Docker
./config-assessment-tool.sh docker --ui
# Or simply (defaults to UI)
./config-assessment-tool.sh docker

# From Source
./config-assessment-tool.sh --ui
```

**Run Headless (no Web UI):** 
   - Run the assessment directly without the UI. This will run CAT and output the generated report files to the output directory and terminate once done. 
```bash
# Using Docker
./config-assessment-tool.sh docker -j <job-file-name>   # dont include the .json extension 
# Example: ./config-assessment-tool.sh docker -j MyJob

# From Source
./config-assessment-tool.sh -j <job-file-name>   # dont include the .json extension
```

## Plugin Management
```bash
# List available plugins
./config-assessment-tool.sh --plugin list

# Start a specific plugin
./config-assessment-tool.sh --plugin start <plugin_name>
```

## Shutdown
Stops running containers and cleanup processes.
```bash
./config-assessment-tool.sh shutdown
```

## Getting Help
Display the help menu with all available arguments.
```bash
./config-assessment-tool.sh --help
```

---

# 4. Configuration (Important)

Job files are configured as below.  There is a default job file provided in the source code at `input/jobs/DefaultJob.json` that you can edit with your controller credentials and settings. You can also create a new job file by copying this default one and editing as needed. If you are using the Web UI, you can also create a job file directly from the UI which will save them into the `input/jobs` directory using the hostname as the job file name however if subsequent edits to the file is needed, you will need to edit the file directly.

Edit `input/jobs/DefaultJob.json` (or your custom job file):

```json
{
  "host": "acme.saas.appdynamics.com",
  "port": 443, 
  "account": "acme",  
  "authType": "token", 
  "username": "api-client-name", 
  "pwd": "temporary-access-token", 
  "verifySsl": true,
  "useProxy": true,
  "applicationFilter": {
    "apm": ".*",
    "mrum": ".*",
    "brum": ".*"
  },
  "timeRangeMins": 1440
}
```

## Authentication Options (`authType`)
*   **`basic`**: Use Controller UI username and password. (Legacy)
*   **`token`** (Recommended):
    *   `username`: Use the API Client Name.
    *   `pwd`: Use the Temporary Access Token.
*   **`secret`**:
    *   `username`: Use the API Client Name.
    *   `pwd`: Use the Client Secret.
