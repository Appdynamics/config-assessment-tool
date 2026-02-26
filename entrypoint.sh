#!/bin/bash

# If no arguments are provided, or help is requested, print usage instructions
if [ $# -eq 0 ] || [ "$1" == "--help" ] || [ "$1" == "-h" ] || [ "$1" == "help" ]; then
  echo ""
  echo "Usage: docker run [DOCKER_OPTIONS] <docker-image> [ --ui | ARGS ]"
  echo ""
  echo "docker container requires you to provide information on where to "
  echo "look for input jobs as well as output and log directories using "
  echo "the -v option to mount local directories into the container."
  echo " "
  echo "DOCKER_OPTIONS:"
  echo "  -p 8501:8501                        What port to access the Web UI if using --ui option. defaults to 8501"
  echo "  -v <local input dir>:/app/input     Required. Must contain 'jobs' and 'thresholds' subfolders."
  echo "  -v <local output dir>:/app/output   Required. Destination for generated reports and archive dir."
  echo "  -v <local logs dir>:/app/logs       Recommended. Where to log job run output"
  echo ""
  echo "Example directory structure required:"
  echo ""
  echo "  configuration-assessment-tool/"
  echo "  ├── input"
  echo "  │   ├── jobs"
  echo "  │   │   └── DefaultJob.json           # Job file (can have multiple job files)"
  echo "  │   └── thresholds"
  echo "  │       └── DefaultThresholds.json    # Default thresholds file is good for most use cases"
  echo "  ├── logs                              # Where CAT logs program output. Optional. Created if not provided"
  echo "  └── output                            # Where all reports are saved. Required. Created if not provided"
  echo ""
  echo "  --ui              Start the Web UI"
  echo "  [ARGS]            Start the Backend (Headless) without UI (see below for options)"
  echo ""
  echo "[ARGS]:"
  echo "  -j, --job-file <name>               Job file name (default: DefaultJob)"
  echo "  -t, --thresholds-file <name>        Thresholds file name (default: DefaultThresholds)"
  echo "  -d, --debug                         Enable debug logging"
  echo "  -c, --concurrent-connections <n>    Number of concurrent connections"
  echo ""
  echo "Examples:"
  echo "1. Run the Web UI: assuming you are running docker using above example directory structure"
  echo "   docker run --rm -p 8501:8501 \\"
  echo "     -v \$(pwd)/input:/app/input \\"
  echo "     -v \$(pwd)/output:/app/output \\"
  echo "     -v \$(pwd)/logs:/app/logs \\"
  echo "     <docker-image> --ui"
  echo ""
  echo "2. Run the Backend (Headless):"
  echo "   docker run --rm -p 8501:8501 \\"
  echo "     -v \$(pwd)/input:/app/input \\"
  echo "     -v \$(pwd)/output:/app/output \\"
  echo "     -v \$(pwd)/logs:/app/logs \\"
  echo "     <docker-image> -j <job-file>"
  echo ""
  echo "for more help run:"
  echo "   docker run <docker-image> --help"
  exit 0
fi

# Fix ModuleNotFoundError: Add current dir (/app) to path
# This allows 'import backend' to work from bin/bundle_main.py export PYTHONPATH=$PYTHONPATH:$(pwd)

# Delegate ALL argument handling to bundle_main.py
# This script handles --ui, --run, --help, and backend [options]
exec python bin/bundle_main.py "$@"
