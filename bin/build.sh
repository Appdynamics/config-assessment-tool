#!/bin/bash

SCRIPT_DIR=$(dirname "$0")
cd "$SCRIPT_DIR" && cd ..|| exit

# Script will build the Docker images from source or pull from DockerHub
if [ -f backend/Dockerfile ] && [ -f frontend/Dockerfile ]; then
    echo "Building appdynamics/config-assessment-tool-backend:latest from Dockerfile"
    docker build -t appdynamics/config-assessment-tool-backend:latest -f backend/Dockerfile .
    echo "Building appdynamics/config-assessment-tool-frontend:latest from Dockerfile"
    docker build -t appdynamics/config-assessment-tool-frontend:latest -f frontend/Dockerfile .
else
    echo "Dockerfiles not found in either backend/ or frontend/."
fi