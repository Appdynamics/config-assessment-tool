# Demo Hybrid Plugin

This plugin demonstrates the "Hybrid" capability, allowing a plugin to function both as an integrated step in the assessment lifecycle and as a standalone CLI tool.

## Overview

- **Name**: `demo_integrated_plugin`
- **Type**: Integrated 
- **Entry Point**: `run_plugin(context)`

## Features

1.  **Integrated plugin**:
    -   Contains a `run_plugin(context)` function.
    -   Automatically executed by the Engine at the end of an assessment job.
    -   Receives the job context (e.g., job file name, output directories).

## How to Run

**Integrated:**
Runs automatically after a job completes.
```bash
./config-assessment-tool.sh --start -j DefaultJob
```
