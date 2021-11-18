#!/bin/bash

SCRIPT_DIR=$(dirname "$0")

cd "$SCRIPT_DIR" && cd ..|| exit

zip -r bin/config-assessment-tool.zip \
 README.md \
 bin/pull.sh \
 bin/run.sh \
 input/jobs/DefaultJob.json \
 input/thresholds/DefaultThresholds.json \
 frontend/FileHandler.py
