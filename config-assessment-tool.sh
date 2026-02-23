#!/bin/bash

OS=$(uname -s | tr '[:upper:]' '[:lower:]')
ARCH=$(uname -m)

if [[ "$OS" == "darwin" ]]; then
  OS_TAG="macos"
elif [[ "$OS" == "linux" ]]; then
  OS_TAG="linux"
else
  echo "Unsupported OS: $OS"
  exit 1
fi

if [[ "$ARCH" == "arm64" || "$ARCH" == "aarch64" ]]; then
  ARCH_TAG="arm"
elif [[ "$ARCH" == "x86_64" ]]; then
  ARCH_TAG="amd64"
else
  echo "Unsupported architecture: $ARCH"
  exit 1
fi

# Detect Repo Owner from Git
if command -v git &> /dev/null; then
  GIT_ORIGIN=$(git config --get remote.origin.url)
  # Extract username from git@github.com:user/repo.git or https://github.com/user/repo.git
  if [[ "$GIT_ORIGIN" =~ github.com[:/]([^/]+)/ ]]; then
    DETECTED_OWNER="${BASH_REMATCH[1]}"
  fi
fi

# Use appdynamics if no owner detected, or honor CAT_REPO_OWNER env var
if [[ -n "$CAT_REPO_OWNER" ]]; then
  REPO_OWNER="$CAT_REPO_OWNER"
elif [[ -z "$DETECTED_OWNER" ]]; then
  REPO_OWNER="appdynamics"
else
  REPO_OWNER="$DETECTED_OWNER"
fi

# Convert to lowercase to be safe for docker images
REPO_OWNER=$(echo "$REPO_OWNER" | tr '[:upper:]' '[:lower:]')

# We now use a single multi-arch manifest image
IMAGE_NAME="ghcr.io/${REPO_OWNER}/config-assessment-tool"

if [[ -f VERSION ]]; then
  VERSION=$(cat VERSION)
else
  echo "VERSION file not found."
  exit 1
fi

# Ensure checking/retagging happens before usage
# If user has the v-prefixed image locally but we want standard no-v image, retag it.
V_VERSION="v$VERSION"
V_IMAGE="$IMAGE_NAME:$V_VERSION"
NO_V_IMAGE="$IMAGE_NAME:$VERSION"

if docker image inspect "$V_IMAGE" >/dev/null 2>&1; then
    if ! docker image inspect "$NO_V_IMAGE" >/dev/null 2>&1; then
        echo "Found image $V_IMAGE, retagging to $NO_V_IMAGE for consistency..."
        docker tag "$V_IMAGE" "$NO_V_IMAGE"
    fi
fi

check_available_images() {
  echo "  Checking available Docker images..."
  local found_images=false
  local repo_url="https://github.com/${REPO_OWNER}?tab=packages&repo_name=config-assessment-tool"

  # Check if curl is available
  if ! command -v curl &> /dev/null; then
    echo "    (curl not found, cannot automatically verify images)"
    echo "    Please check the repository manually for available packages at:"
    echo "    $repo_url"
    return
  fi

  # Query GitHub Packages page for this user/org
  local html_content
  if ! html_content=$(curl -s -L "$repo_url"); then
     echo "    (Failed to connect to GitHub to verify images)"
     echo "    Please check the repository manually for available packages at:"
     echo "    $repo_url"
     return
  fi

  echo "  Available Images (verified on GitHub for user '$REPO_OWNER'):"

  # Check for the main package which now holds multi-arch manifest
  if echo "$html_content" | grep -q "config-assessment-tool"; then
    echo "    $IMAGE_NAME:$VERSION (Multi-arch: Linux x86/ARM/RISC-V, Windows x86)"
    found_images=true
  else
    # Fallback to checking for specific platform tags if main one isn't clear from HTML grep
    # (HTML grep is brittle, but requested method)
    # Check Linux Multiarch
    if echo "$html_content" | grep -q "config-assessment-tool-linux-multiarch"; then
      echo "    $IMAGE_NAME-linux-multiarch:$VERSION"
      found_images=true
    fi
    # Check Windows
    if echo "$html_content" | grep -q "config-assessment-tool-windows-amd64"; then
      echo "    $IMAGE_NAME-windows-amd64:$VERSION"
      found_images=true
    fi
  fi

  if [ "$found_images" = false ]; then
    echo "    (No verified images found locally for version $VERSION)"
    # Fallback to listing expected ones if verify fails or none found?
    echo "    $IMAGE_NAME:$VERSION (Linux/Windows Multi-arch)"
  fi
  echo "  "
}

IMAGE="$IMAGE_NAME:$VERSION"
PORT="8501"
LOG_DIR="logs"
LOG_FILE="$LOG_DIR/config-assessment-tool.log"
MOUNTS="-v $(pwd)/input/jobs:/app/input/jobs -v $(pwd)/input/thresholds:/app/input/thresholds -v $(pwd)/output/archive:/app/output/archive -v $(pwd)/$LOG_DIR:/app/$LOG_DIR"
CONTAINER_NAME="cat-tool-container"

mkdir -p "$LOG_DIR"

start_filehandler() {
  if [ ! -f "frontend/FileHandler.py" ]; then
    echo "Error: frontend/FileHandler.py not found."
    exit 1
  fi
  echo "Starting FileHandler service on host..."
  pkill -f "python.*FileHandler.py" 2>/dev/null
  pipenv run python frontend/FileHandler.py >> "$LOG_FILE" 2>&1 &
  echo "FileHandler.py started with PID $!"
  sleep 2
}

# Helper function to run docker logic
run_docker() {
    export FILE_HANDLER_HOST=host.docker.internal
    start_filehandler
    docker stop $CONTAINER_NAME >/dev/null 2>&1
    docker rm $CONTAINER_NAME >/dev/null 2>&1

    # Check if UI Mode requested (No args, or --ui/--run flags in first position after 'docker')
    # $1 is the first arg after 'docker'
    IS_UI_MODE=false
    if [[ $# -eq 0 ]]; then
        IS_UI_MODE=true
    elif [[ "$1" == "--ui" || "$1" == "--run" ]]; then
        IS_UI_MODE=true
    fi

    if [[ "$IS_UI_MODE" == "true" ]]; then
      echo "Starting container in UI mode..."
      # Pass --ui explicitly to trigger UI mode in entrypoint
      CONTAINER_ID=$(docker run --add-host=host.docker.internal:host-gateway -d --name $CONTAINER_NAME -e FILE_HANDLER_HOST=$FILE_HANDLER_HOST -p $PORT:$PORT $MOUNTS $IMAGE --ui)
      if [ $? -eq 0 ]; then
        echo "Container started successfully with ID: $CONTAINER_ID"
        echo "UI available at http://localhost:$PORT"
        docker logs -f $CONTAINER_ID
      else
        echo "Failed to start container."
        exit 1
      fi
    else
      echo "Starting container in backend mode with args: $@"
      # Pass arguments directly without prepending 'backend'
      docker run --add-host=host.docker.internal:host-gateway --rm --name $CONTAINER_NAME -e FILE_HANDLER_HOST=$FILE_HANDLER_HOST -p $PORT:$PORT $MOUNTS $IMAGE "$@"
      EXIT_CODE=$?
      if [ $EXIT_CODE -ne 0 ]; then
        echo "Container failed with exit code: $EXIT_CODE"
        exit $EXIT_CODE
      fi
    fi
}

# Helper function to run source logic
run_source() {
    export PYTHONPATH="$(pwd):$(pwd)/backend"

    if ! command -v pipenv &> /dev/null; then
      echo "pipenv not found. Attempting to install via pip..."
      pip install pipenv
    fi

    # Ensure dependencies are installed before running
    echo "Checking/Installing dependencies..."
    pipenv install

    # Unified Argument Handling for Source Mode
    # We check if the first arg is --ui or --run to launch UI, otherwise backend.
    # Default to UI if no args are provided ($1 is empty)
    if [[ -z "$1" || "$1" == "--ui" || "$1" == "--run" ]]; then
      echo "PYTHONPATH is: $PYTHONPATH"
      echo "Running application in UI mode from source..."

      echo "UI available at http://localhost:$PORT"
      pipenv run streamlit run frontend/frontend.py
    else
      echo "PYTHONPATH is: $PYTHONPATH"
      echo "Running application in backend mode from source with args: $@"
      pipenv run python backend/backend.py "$@"
    fi
}

case "$1" in
  docker)
    shift
    run_docker "$@"
    ;;

  --plugin)
    if [[ "$2" == "list" ]]; then
       export PYTHONPATH="$(pwd):$(pwd)/backend"
       pipenv run python backend/plugin_manager.py list
       exit 0
    elif [[ "$2" == "docs" ]]; then
       PLUGIN_NAME="$3"
       if [[ -z "$PLUGIN_NAME" ]]; then
         echo "Error: Plugin name required."
         exit 1
       fi
       export PYTHONPATH="$(pwd):$(pwd)/backend"
       pipenv run python backend/plugin_manager.py docs "$PLUGIN_NAME"
       exit 0
    elif [[ "$2" == "start" ]]; then
       PLUGIN_NAME="$3"
       if [[ -z "$PLUGIN_NAME" ]]; then
         echo "Error: Plugin name required."
         exit 1
       fi
       export PYTHONPATH="$(pwd):$(pwd)/backend"
       # Pass remaining args to the plugin manager
       pipenv run python backend/plugin_manager.py start "$PLUGIN_NAME" "${@:4}"
       exit 0
    fi
    ;;

  shutdown)
    echo "Shutting down container: $CONTAINER_NAME"
    docker stop $CONTAINER_NAME >/dev/null 2>&1
    docker rm $CONTAINER_NAME >/dev/null 2>&1
    echo "Container stopped and removed."
    echo "Stopping FileHandler process..."
    pkill -f "python.*FileHandler.py" 2>/dev/null
    echo "FileHandler stopped."
    echo "Stopping backend process..."
    pkill -f "python.*backend.py" 2>/dev/null
    echo "Backend process stopped."
    echo "Stopping Streamlit process..."
    pkill -f "streamlit run frontend/frontend.py" 2>/dev/null
    echo "Streamlit stopped."
    ;;

  --help|help|"")
    echo "Usage:"
    echo "  config-assessment-tool [--ui]                   # Starts CAT UI from source (Default). Requires Python 3.12 and pipenv."
    echo "  config-assessment-tool [OPTIONS]                # Starts CAT headless mode from source with [OPTIONS]."
    echo "  config-assessment-tool docker [--ui]            # Starts CAT UI using Docker. Docker install required."
    echo "  config-assessment-tool docker [OPTIONS]         # Starts CAT headless mode using Docker with [OPTIONS]. Docker required."
    echo "  config-assessment-tool --plugin <list|start|docs> [name]    # Manage plugins"
    echo "  config-assessment-tool shutdown                 # Stop and remove the running container and FileHandler"
    echo ""
    echo "[OPTIONS]:"
    echo "  -j, --job-file <name>             Job file name (default: DefaultJob)"
    echo "  -t, --thresholds-file <name>      Thresholds file name (default: DefaultThresholds)"
    echo "  -d, --debug                       Enable debug logging"
    echo "  -c, --concurrent-connections <n>  Number of concurrent connections"
    echo "  "
    echo "Direct Docker Usage:"
    echo "  You can also run the tool directly using Docker without this script."
    echo "  Ensure you mount the input, output, and logs directories."
    echo "  "
    echo "  # UI Mode (Auto-detects architecture: x86/ARM/RISC-V):"
    echo "  docker run -p 8501:8501 \\"
    echo "    -v \$(pwd)/input:/app/input \\"
    echo "    -v \$(pwd)/output:/app/output \\"
    echo "    -v \$(pwd)/logs:/app/logs \\"
    echo "    $IMAGE_NAME:$VERSION --ui"
    echo "  "
    echo "  # Headless Mode:"
    echo "  docker run --rm \\"
    echo "    -v \$(pwd)/input:/app/input \\"
    echo "    -v \$(pwd)/output:/app/output \\"
    echo "    -v \$(pwd)/logs:/app/logs \\"
    echo "    $IMAGE_NAME:$VERSION [OPTIONS]"
    echo "    e.g. $IMAGE_NAME:$VERSION -j <job-file>"
    echo "  "

    check_available_images
    exit 0
    ;;

  *)
    # If it's not a known command or docker, assume it is arguments for running from source
    run_source "$@"
    ;;
esac

