# Plugin Development Guide

This directory allows you to extend the Config Assessment Tool with custom functionality.

## Plugin Structure

Every plugin must be a **directory** containing at least a `main.py` file.

```
plugins/
  ├── my_awesome_plugin/
  │   ├── main.py        <-- Entry point (Required)
  │   ├── utils.py       <-- Helper files (Optional)
  │   └── config.json    <-- Resources (Optional)
```

## Plugin Types

There are two ways a plugin can behave, determined by the code inside `main.py`.

### 1. Integrated Plugin (config-assessment-tool Lifecycle)
These plugins run automatically at the end of every config-assessment-tool job (after reports are generated).

*   **Requirement:** Define a `run_plugin(context)` function.
*   **Context:** Receives a dictionary containing `jobFileName`, `outputDir`, and `controllerData`.

**Example `main.py`:**
```python
import logging

def run_plugin(context):
    """
    Called automatically by the Engine after config-assessment-tool execution.
    """
    job_name = context.get('jobFileName')
    logging.info(f"My Plugin: processing results for {job_name}")
    
    # Return value is logged but not used
    return "Success"
```

### 2. Standalone Plugin (CLI Tool)
These plugins are run manually via the command line and do not interfere with the config-assessment-tool process.

*   **Requirement:** Do **NOT** define a `run_plugin` function.
*   **Execution:** logic runs under `if __name__ == "__main__":`.

**Example `main.py`:**
```python
import sys

def main():
    print("I am a standalone tool!")
    print(f"I received args: {sys.argv[1:]}")

if __name__ == "__main__":
    main()
```

### 3. Hybrid Plugin
A plugin can support both modes!

*   Define `run_plugin(context)` for the engine.
*   Add `if __name__ == "__main__":` block to run purely via CLI (usually passing a dummy context or running a different mode like a UI).

## How to Run

**List available plugins:**
```bash
./config-assessment-tool.sh --plugin list
```

**Run a specific plugin (Standalone Mode):**
```bash
./config-assessment-tool.sh --plugin start <plugin_name> [arg1] [arg2] ...
```

**Run Integrated Plugins:**
Simply run a standard job. All valid plugins found in this folder will execute at the end.
```bash
./config-assessment-tool.sh --start -j DefaultJob
```
