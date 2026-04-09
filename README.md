# Config Assessment Tool

This document provides a quick-start guide for the 1.8.0 release. This release introduces docker usage improvements, UI improvements, feature and bug fixes.

## Table of Contents
1. [Overview](#overview)
2. [Quick start](#quick-start)
3. [Output artifacts](#output-artifacts)
4. [Requirements and limitations](#requirements-and-limitations)
5. [Support](#support)
6. [Troubleshooting](#troubleshooting)

---

## Overview

Config Assessment Tool helps evaluate AppDynamics instrumentation and configuration health, then generates report artifacts your team can use for analysis and follow-up actions. For a more detailed tool overview and architecture, see [`General Description`](docs/OVERVIEW.md#general-description).

---

## Quick start

1. Review prerequisites and prepare your job file in [`docs/RUNBOOK.md`](docs/RUNBOOK.md#prerequisites-and-setup).
2. Choose how you want to run CAT:
   - [Platform executable bundle (easiest way to run)](#platform-executable-easiest-way-to-run)
   - [Docker](#docker)
   - Other ways to run CAT: [`docs/RUNBOOK.md#other-ways-to-run`](docs/RUNBOOK.md#other-ways-to-run)
3. Run the tool and review generated files in `output/`. This directory generates once you run the tool.
4. If you hit issues, use [`docs/RUNBOOK.md`](docs/RUNBOOK.md) and [`docs/TROUBLESHOOTING.md`](docs/TROUBLESHOOTING.md).

### Platform executable (easiest way to run)

1. Download the executable bundle for your platform from [GitHub releases](https://github.com/appdynamics/config-assessment-tool/releases/latest).
2. Extract the bundle.
3. Edit `input/jobs/DefaultJob.json` (or add your own job file in `input/jobs/`).
4. Run one of the commands below from the extracted bundle directory.

macOS/Linux:

```bash
./config-assessment-tool --ui           # run using Web UI mode
./config-assessment-tool -j DefaultJob  # run in headless mode using DefaultJob.json job file
./config-assessment-tool --help         # get CLI help
```

Windows PowerShell:

```powershell
.\config-assessment-tool.exe --ui           # run using Web UI mode
.\config-assessment-tool.exe -j DefaultJob  # run in headless mode using DefaultJob.json job file
.\config-assessment-tool.exe --help         # get CLI help
```

If macOS blocks the executable after extraction run the following command from the extracted bundle directory to remove the quarantine attribute:

```bash
sudo xattr -rd com.apple.quarantine .
```

Full details: [`docs/RUNBOOK.md#run-using-platform-executables`](docs/RUNBOOK.md#run-using-platform-executables)

### Docker

1. Ensure Docker Desktop / Docker Engine is running.
2. From your project directory, confirm local folders exist: `input/`, `output/`, and `logs/`.
3. Ensure your job file exists at `input/jobs/DefaultJob.json`.
4. Run one of the commands below.

macOS/Linux (UI mode):

```bash
docker run \
  --name "config-assessment-tool" \
  -v <config-assessment-tool-directory>/input:/app/input \
  -v <config-assessment-tool-directory>/output:/app/output \
  -v <config-assessment-tool-directory>/logs:/app/logs \
  -p 8501:8501 \
  --rm \
  ghcr.io/appdynamics/config-assessment-tool:latest --ui
```

macOS/Linux (headless run):

```bash
docker run \
  --name "config-assessment-tool" \
  -v <config-assessment-tool-directory>/input:/app/input \
  -v <config-assessment-tool-directory>/output:/app/output \
  -v <config-assessment-tool-directory>/logs:/app/logs \
  --rm \
  ghcr.io/appdynamics/config-assessment-tool:latest -j DefaultJob -t DefaultThresholds
```

Windows PowerShell (UI mode):

```powershell
docker run `
  --name "config-assessment-tool" `
  -v <config-assessment-tool-directory>/input:/app/input `
  -v <config-assessment-tool-directory>/output:/app/output `
  -v <config-assessment-tool-directory>/logs:/app/logs `
  -p 8501:8501 `
  --rm `
  ghcr.io/appdynamics/config-assessment-tool:latest --ui
```

Get CLI help:

```bash
docker run ghcr.io/appdynamics/config-assessment-tool:latest --help
```

Full details: [`docs/RUNBOOK.md#run-using-docker`](docs/RUNBOOK.md#run-using-docker)

## Output artifacts
Generated in `output/{jobName}` (varies by job name):
- `{jobName}-cx-presentation.pptx` # PowerPoint summary report
- `{jobName}-MaturityAssessment-apm.xlsx`  # Summary Excel report for APM
- `{jobName}-MaturityAssessment-brum.xlsx` # Summary Excel report for Browser RUM
- `{jobName}-MaturityAssessment-mrum.xlsx` # Summary Excel report for Mobile RUM
- `{jobName}-AgentMatrix.xlsx`
- `{jobName}-CustomMetrics.xlsx`
- `{jobName}-License.xlsx`  # License usage summary report
- `{jobName}-MaturityAssessmentRaw-apm.xlsx`
- `{jobName}-MaturityAssessmentRaw-brum.xlsx`
- `{jobName}-MaturityAssessmentRaw-mrum.xlsx`
- `{jobName}-ConfigurationAnalysisReport.xlsx` # Prescribed steps to raise maturity levels
- `controllerData.json` # Raw data dump of all controller API responses for debugging and custom analysis
- `info.json`

Generated in `output/archive` directory
- Archived reports organized by timestamp and job name for record-keeping and trend analysis. Every time you run CAT, the output files are also copied to the archive directory with a timestamp and maintained for future reference and analysis.  

---

## Requirements and limitations
Requirements:
- Python 3.12 for source-based runs
- Docker engine for Docker-based runs
- No local Python or Docker required for platform executable bundles

Known limitation:
- Certain data collector snapshot lookups are limited by product/API behavior.

---

## Support
- Open GitHub issues for bugs, feature requests, and feedback:
  `https://github.com/Appdynamics/config-assessment-tool`
- You may include log output snippets, stack trace etc. DO NOT include proprietary data such as controller URL's etc.
- Enable debug via UI checkbox or CLI flags `--debug` / `-d`.

---

## Troubleshooting
For common errors, diagnostics, and fixes, see [`docs/TROUBLESHOOTING.md`](docs/TROUBLESHOOTING.md), including [`Proxy issues`](docs/TROUBLESHOOTING.md#proxy-issues).
