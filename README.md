![Tests](https://github.com/appdynamics/config-assessment-tool/actions/workflows/tests.yml/badge.svg)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
![cod cov](https://github.com/appdynamics/config-assessment-tool/blob/github-workflow/coverage.svg)

# config-assessment-tool

This project aims to provide a single source of truth for performing AppDynamics Health Checks.

## Usage

There are four options to run the tool:

1. UI Method: The config-assessment-tool provides a frontend UI to view/run jobs
2. Platform executable: An Operating System specific bundle if you are not using Docker and Python
3. Directly via Docker: The backend container can be run manually from the command line
4. From Source: Manually install dependencies and run the `backend.py` script directly

### Important step for running on Windows(Ignore this step if using method 2 above - Platform executable)

Docker on Windows requires manually sharing the `/input`, `/output`, and `/logs` directories with the container. If you do not follow this step, you will get the following error when trying to run the
container: `DockerException Filesharing has been cancelled`. Take a look at the documentation [here](https://docs.docker.com/desktop/windows/) for more information.

### Expected Permissions
The tool expects ONLY the following permissions to be given:

- Account Owner (Default)
- Administrator (Default)
- Analytics Administrator (Default)

### UI method

Obtain frontend and backend Docker images via:

1. Download the latest `config-assessment-tool.zip` from [here](https://github.com/Appdynamics/config-assessment-tool/releases)
2. Pull from ghrc with `python3 bin/config-assessment-tool.py --pull`
3. Run with `python3 bin/config-assessment-tool.py --run`
4. Navigate to `http://localhost:8501`

![Scheme](frontend/resources/img/frontend.png)

Add new Jobs or Thresholds to `config_assessment_tool/resources/jobs` and `config_assessment_tool/resources/thresholds` respectively.

Refresh the page to see the Jobs and Thresholds appear.

### Platform executable

Use this method if you are not able to use Docker or Python in your target deployment environment. Currently, platform bundles are available for Windows and Linux only.

1. Download and unzip (or untar in case of linux) the latest `config-assessment-tool-<OS>-<version>.<zip|tgz>` from [here](https://github.com/Appdynamics/config-assessment-tool/releases) where OS is one of windows/linux depending on your target host and version is the config tool release version
2. cd into the expanded directory and edit `input/jobs/DefaultJobs.json` to match your target controller. You may also create a job file of your own. e.g. `input/jobs/<job-file-name>.json`
3. Run the executable for your target platform located at the root of expanded directory:
1. For Linux: using a command line shell/terminal run `./config-assessment-tool` if using DefaultJob.json or `./config-assessment-tool -j <job-file-name>` if you created your own job file
2. For Windows: using a CMD or PowerShell terminal run `.\config-assessment-tool.exe` if using DefaultJob.json or `./config-assessment-tool.exe -j <job-file-name>` if you created your own job file

This method of running the tool currently does not support using the UI. You may only use command line instructions as outlined above. You can change the settings by editing the included configuration files directly.  You may ignore any other files/libraries in the bundle. The configuration files and their directory locations for you to edit are listed below.
```

config-assessment-tool-<OS>-<version>/
├── config-assessment-tool          # executable file to run. For Windows this will be config-assessment-tool.exe
├── input
│   ├── jobs
│   │   └── DefaultJob.json         # default job used if no job file flag used with your own custom job file(-j). Your Controller(s) connection settings.
│   └── thresholds
│       └── DefaultThresholds.json  # default threshold file used if no custom threshhold file flag(-t) is used with your own custom file 
│   ├── ....
│   │
└── ...
    └── ...

```

### Directly via Docker

You can start the backend container with the following command:

Unix

```
docker run \
--name "config-assessment-tool-backend" \
-v "$(pwd)/logs":/logs \
-v "$(pwd)/output":/output \
-v "$(pwd)/input":/input \
-e HOST_ROOT="$(pwd)" \
-p 8501:8501 \
--rm \
ghcr.io/appdynamics/config-assessment-tool-backend:latest -j acme -t DefaultThresholds
```

Windows

```
docker run `
--name "config-assessment-tool-backend" `
-v $pwd/logs:/logs `
-v $pwd/output:/output `
-v $pwd/input:/input `
-e HOST_ROOT=$pwd `
-p 8501:8501 `
--rm `
ghcr.io/appdynamics/config-assessment-tool-backend:latest -j acme -t DefaultThresholds
```

### From Source

#### Steps to run

Required

1. `git clone https://github.com/Appdynamics/config-assessment-tool.git`
2. `cd config-assessment-tool`
3. `pipenv install`
4. `pipenv shell`
5. `python3 backend/backend.py -j acme -t DefaultThresholds`

```
Usage: backend.py [OPTIONS]

Options:
  -j, --job-file TEXT
  -t, --thresholds-file TEXT
  -d, --debug
  -c, --concurrent-connections INTEGER
  --help                          Show this message and exit.
```

Options `--job-file` and `--thresholds-file` will default to `DefaultJob` and `DefaultThresholds` respectively.

All Job and Threshold files must be contained in `config_assessment_tool/resources/jobs` and `config_assessment_tool/resources/thresholds` respectively. They are to be referenced by name file name (
excluding .json), not full path.

The frontend can be invoked by navigating to `config_assessment_tool/frontend` and invoking `streamlit run frontend.py`

## Output

This program will create the following files in the `out` directory.

- `{jobName}-BSGReport-apm.xlsx`
  - Bronze/Silver/Gold report
- `{jobName}-Agent-Matrix.xlsx`
  - Details agent versions rolled up by application
  - Lists the details of individual without any rollup
- `{jobName}-CustomMetricsReport.xlsx`
  - Lists which applications are leveraging Custom Extensions
- `{jobName}-License.xlsx`
  - Export of the License Usage page in the Controller
- `{jobName}-RawBSGReport.xlsx`
  - Raw metrics which go into the above BSG report
- `controllerData.json`
  - Contains all raw data used in analysis.
- `info.json`
  - Contains information on previous job execution.

## Program Architecture

### Backend

![Scheme](backend/resources/img/architecture.jpg)

## Proxy Support

Support for plain HTTP proxies and HTTP proxies that can be upgraded to HTTPS via the HTTP CONNECT method is provided by enabling the `useProxy` flag in a given job file. Enabling this flag will cause
the backend to use the proxy specified from environment variables: HTTP_PROXY, HTTPS_PROXY, WS_PROXY or WSS_PROXY (all are case insensitive). Proxy credentials are given from ~/.netrc file if present.
See aiohttp.ClientSession [documentation](https://docs.aiohttp.org/en/stable/client_advanced.html#proxy-support) for more details.

## Requirements

- Python 3.5 or above if running with `bin/config-assessment-tool.py`
- Python 3.9 or above if running from source
- [Docker](https://www.docker.com/products/docker-desktop)
- None if running using Platform executable method. Tested on most Linux distributions and Windows 10/11

## Limitations

- Data Collectors
  - The API to directly find snapshots containing data collectors of type `Session Key` or `HTTP Header` does not work.
  - The API does however work for `Business Data` (POJO match rule), `HTTP Parameter`, and `Cookie` types.
  - As far as I can tell this is a product limitation, the transaction snapshot filtering UI does not even have an option for `Session Key` or `HTTP Header`.
  - The only way to check for `Session Key` or `HTTP Header` data collector existence within snapshots would be to inspect ALL snapshots (prohibitively time intensive).
  - As a workaround, we will assume any `Session Key` or `HTTP Header` data collectors are present in snapshots.

## Support

Please email bradley.hjelmar@appdynamics.com for any issues and attach debug logs.

Debug logs can be taken by either:

- checking the `debug` checkbox in the UI
- running the backend with the `--debug` or `-d` flag