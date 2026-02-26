# Demo Hybrid Plugin

This plugin demonstrates the "Hybrid" capability, allowing a plugin to function both as an integrated step in the assessment lifecycle and as a standalone CLI tool.

## Overview

- **Name**: `demo_hybrid`
- **Type**: Hybrid (Integrated + Standalone)
- **Entry Point**: `main.py`

## Features

1.  **Integrated Mode**:
    -   Contains a `run_plugin(context)` function.
    -   Automatically executed by the Engine at the end of an assessment job.
    -   Receives the job context (e.g., job file name, output directories).

2.  **Standalone Mode**:
    -   Contains an `if __name__ == "__main__":` block.
    -   Can be run manually via the CLI.
    -   Useful for testing logic independently or providing utility functions.

## How to Run

**Standalone Mode (CLI):**
Runs independently with dummy data.
```bash
./config-assessment-tool.sh --plugin start demo_hybrid
```

**Integrated Mode (Assessment):**
Runs automatically after a job completes.
```bash
./config-assessment-tool.sh --start -j DefaultJob
```
