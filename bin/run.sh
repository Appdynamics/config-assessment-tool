#!/bin/bash

cat << EOF
                  __ _                                                              _        _              _
  ___ ___  _ __  / _(_) __ _        __ _ ___ ___  ___  ___ ___ _ __ ___   ___ _ __ | |_     | |_ ___   ___ | |
 / __/ _ \| '_ \| |_| |/ _\` |_____ / _\` / __/ __|/ _ \/ __/ __| '_ \` _ \ / _ \ '_ \| __|____| __/ _ \ / _ \| |
| (_| (_) | | | |  _| | (_| |_____| (_| \__ \__ \  __/\__ \__ \ | | | | |  __/ | | | ||_____| || (_) | (_) | |
 \___\___/|_| |_|_| |_|\__, |      \__,_|___/___/\___||___/___/_| |_| |_|\___|_| |_|\__|     \__\___/ \___/|_|
                       |___/
EOF

PYTHON3_REF=$(which python3 | grep "/python3")
PYTHON_REF=$(which python | grep "/python")

if [ -z "$PYTHON3_REF" ] && [ -z "$PYTHON_REF" ]; then
  echo Python not found. Install python to run the script.
  return 1 2> /dev/null || exit
fi

SCRIPT_DIR=$(dirname "$0")

cd "$SCRIPT_DIR" && cd .. || exit

mkdir -p "$(pwd)/logs"
mkdir -p "$(pwd)/output"

# Kill FileHandler if already running
FILE_HANDLER_PID=$(pgrep -f 'frontend/FileHandler.py')
if [ -n "$FILE_HANDLER_PID" ]; then
  kill "$FILE_HANDLER_PID"
fi

# Kill config-assessment-tool-frontend if already running
dsi() {
  PID=$(docker ps | awk -v i="^$1.*" '{if($2~i){print$1}}');
  if [ -n "$PID" ]; then
    echo "Stopping config-assessment-tool-frontend"
    docker stop "$PID" 1> /dev/null 2> /dev/null
  fi
}
dsi "ghcr.io/appdynamics/config-assessment-tool-frontend:latest"

# start up our FileHandler server locally, will exit when this process exits
trap "kill 0" SIGINT
if [ -n "$PYTHON3_REF" ]; then
  python3 -u frontend/FileHandler.py >>./logs/config-assessment-tool-frontend.log &
else
  python -u frontend/FileHandler.py >>./logs/config-assessment-tool-frontend.log &
fi

# 8501 is used by the frontend Streamlit app
# 1337 is used by the FileHandler server

docker run \
  --name "config-assessment-tool-frontend" \
  -v /var/run/docker.sock:/var/run/docker.sock \
  -v "$(pwd)/logs":/logs \
  -v "$(pwd)/output":/output \
  -v "$(pwd)/input":/input \
  -e HOST_ROOT="$(pwd)" \
  -p 8501:8501 \
  -p 1337:1337 \
  --rm \
  ghcr.io/appdynamics/config-assessment-tool-frontend:latest &

echo "waiting for config-assessment-tool-frontend to start"
while [ "$(docker inspect -f {{.State.Running}} config-assessment-tool-frontend 2> /dev/null)" != "true" ]; do sleep 2; done
echo "config-assessment-tool-frontend is running"

echo "opening browser to http://localhost:8501"
case "$(uname -a)" in
   *Darwin*)
     open http://localhost:8501
     ;;

   *WSL*)
     wslview http://localhost:8501
     ;;

   *Linux*)
     xdg-open http://localhost:8501
     ;;

   *CYGWIN*|*MINGW32*|*MSYS*|*MINGW*)
     Start-Process "http://localhost:8501"
     ;;

   *)
     echo 'Other OS'
     ;;
esac

# idle waiting for abort from user
echo "press ctrl+c to exit"
read -r -d '' _ </dev/tty
