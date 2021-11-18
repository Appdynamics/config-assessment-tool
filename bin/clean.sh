#!/bin/bash

docker images -a | grep "appdynamics/config-assessment-tool-backend" | awk '{print $3}' | xargs docker rmi
docker images -a | grep "appdynamics/config-assessment-tool-frontend" | awk '{print $3}' | xargs docker rmi
