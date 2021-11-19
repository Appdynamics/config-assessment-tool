![Tests](https://github.com/appdynamics/config-assessment-tool/actions/workflows/tests.yml/badge.svg)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
![cod cov](https://github.com/appdynamics/config-assessment-tool/blob/github-workflow/coverage.svg)

# config-assessment-tool

This project aims to provide a single source of truth for performing AppDynamics Health Checks.

## Usage

There are three options to run the tool:

1. The config-assessment-tool provides a frontend UI to view/run jobs. 

2. The container can be run manually from the command line.

3. Build and run the `backend.py` script directly.

### UI method

1. Build docker image dependencies with `./bin/pull.sh`.
2. Run with `./bin/run.sh`.
3. Navigate to `http://localhost:8501`

![Scheme](frontend/resources/img/frontend.png)

Add new Jobs or Thresholds to `config_assessment_tool/resources/jobs` and `config_assessment_tool/resources/thresholds` respectively.

Refresh the page to see the Jobs and Thresholds appear.

### Directly via Docker

You can start the container with the following command:

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

### Local

The backend can be invoked via `python backend.py`.

```
Usage: backend.py [OPTIONS]

Options:
  -j, --job-file FILENAME
  -t, --thresholds-file FILENAME
  --help Show this message and exit.
```

Options `--job-file` and `--thresholds-file` will default to `DefaultJob` and `DefaultThresholds` respectively.

All Job and Threshold files must be contained in `config_assessment_tool/resources/jobs` and `config_assessment_tool/resources/thresholds` respectively.
They are to be referenced by name file name (excluding .json), not full path.

The frontend can be invoked by navigating to `config_assessment_tool/frontend` and invoking `streamlit run frontend.py`

## Output

This program will generate `{jobName}.xlsx` in the `out` directory containing the Health Check analysis.

## Program Architecture

### Backend

![Scheme](backend/resources/img/architecture.jpg)

## Requirements
- Either Python 2 or 3 
- Docker

## Limitations
- Data Collectors
  - The API to directly find snapshots containing data collectors of type `Session Key` or `HTTP Header` does not work.
  - The API does however work for `Business Data` (POJO match rule), `HTTP Parameter`, and `Cookie` types.
  - As far as I can tell this is a product limitation, the transaction snapshot filtering UI does not even have an option for `Session Key` or `HTTP Header`. 
  - The only way to check for `Session Key` or `HTTP Header` data collector existence within snapshots would be to inspect ALL snapshots (prohibitively time intensive).
  - As a workaround, we will assume any `Session Key` or `HTTP Header` data collectors are present in snapshots.


## TODO
- Custom Metrics Report
- User Audit Report
- Tie Users to Dashboards and applications
- 


## Support
Please email bradley.hjelmar@appdynamics.com for any issues.