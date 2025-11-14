import logging
import re
import time
from collections import OrderedDict

from api.appd.AppDService import AppDService
from extractionSteps.JobStepBase import JobStepBase
from util.asyncio_utils import AsyncioUtils
from util.stdlib_utils import substringBetween


class MachineAgentsAPM(JobStepBase):
    def __init__(self):
        super().__init__("apm")

    async def extract(self, controllerData):
        """
        Extract node level details.
        1. Makes one API call per application to get node Machine Agent Availability.
        2. Is dependent on nodes from AppAgents.
        """
        jobStepName = type(self).__name__

        nodeIdToMachineAgentAvailabilityMap = {}
        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Extracting {jobStepName}')
            controller: AppDService = hostInfo["controller"]

            # Gather necessary metrics.
            machineAgentAvailabilityFutures = []
            for application in hostInfo[self.componentType].values():
                machineAgentAvailabilityFutures.append(
                    controller.getMetricData(
                        applicationID=application["id"],
                        metric_path="Application Infrastructure Performance|*|Individual Nodes|*|Agent|Machine|Availability",
                        rollup=True,
                        time_range_type="BEFORE_NOW",
                        duration_in_mins=controller.timeRangeMins,
                    )
                )
            machineAgentAvailability = await AsyncioUtils.gatherWithConcurrency(*machineAgentAvailabilityFutures)

            # Create a dictionary of Node -> Calls Per Minute for fast lookup
            for rolledUpMetrics in machineAgentAvailability:
                if rolledUpMetrics.error is not None:  # call to gather metrics failed for some reason (most likely 504)
                    continue
                for nodeMetric in rolledUpMetrics.data:
                    try:
                        # e.g. 'Overall Application Performance|foo|Individual Nodes|bar|Calls per Minute'
                        tierName = substringBetween(
                            nodeMetric["metricPath"],
                            left="Application Infrastructure Performance|",
                            right="|",
                        )
                        nodeName = substringBetween(
                            nodeMetric["metricPath"],
                            left="Individual Nodes|",
                            right="|",
                        )
                        machineAgentAvailabilityMetric = nodeMetric["metricValues"][0]["sum"]
                    except IndexError:
                        tierName = ""
                        nodeName = ""
                        machineAgentAvailabilityMetric = 0
                    nodeIdToMachineAgentAvailabilityMap[tierName + "|" + nodeName] = machineAgentAvailabilityMetric

            # Append node level information to overall host info
            hostInfo["nodeMachineIdMachineAgentAvailabilityMap"] = {}
            for application in hostInfo[self.componentType]:
                for node in hostInfo[self.componentType][application]["nodes"]:
                    try:
                        node["machineAgentAvailability"] = nodeIdToMachineAgentAvailabilityMap[node["tierName"] + "|" + node["name"]]
                    except (KeyError, TypeError):
                        node["machineAgentAvailability"] = 0
                        logging.debug(
                            f'{hostInfo["controller"].host} - Node: {node["tierName"]}|{node["name"]} returned no metric data for Agent Availability.'
                        )
                    hostInfo["nodeMachineIdMachineAgentAvailabilityMap"][node["machineId"]] = (
                        node["machineAgentAvailability"] / controller.timeRangeMins * 100
                    )

    def analyze(self, controllerData, thresholds):
        """
        Analysis of node level details.
        1. Determines machine agent age from semantic versioning. (Version 4.X and under will always fail).
        2. Determines number of agents reporting data.
        3. Determines number of agents running same version. In the case of multiple versions, will return the largest common agent count regardless of version.
        4. Determines number of App Agent nodes with an installed Machine Agent.
        """

        # Used to determine agent age from semantic versioning of agents
        currYearAndMonth = [int(x) for x in time.strftime("%Y,%m").split(",")]
        currYear = int(str(currYearAndMonth[0])[-2:])
        currMonth = currYearAndMonth[1]

        semanticVersionRegex = re.compile("[0-9]+\\.[0-9]+\\.")

        jobStepName = type(self).__name__

        # Get thresholds related to job
        jobStepThresholds = thresholds[self.componentType][jobStepName]

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Analyzing {jobStepName}')

            hostInfo["machineAgentVersions"] = set()

            for application in hostInfo[self.componentType].values():
                # Root node of current application for current JobStep.
                analysisDataRoot = application[jobStepName] = OrderedDict()
                # This data goes into the 'JobStep - Metrics' xlsx sheet.
                analysisDataEvaluatedMetrics = analysisDataRoot["evaluated"] = OrderedDict()
                # This data goes into the 'JobStep - Raw' xlsx sheet.
                analysisDataRawMetrics = analysisDataRoot["raw"] = OrderedDict()

                numberNodesWithMachineAgentInstalled = 0.0
                numberMachineAgentsLessThan1YearOld = 0
                numberMachineAgentsLessThan2YearsOld = 0
                numberMachineAgentsReportingData = 0
                numberMachineAgentsRunningSameVersion = 0
                numberMachineAgentsInstalledAlongsideAppAgents = 0
                nodeVersionMap = {}

                application["machineAgentVersions"] = []

                for node in application["nodes"]:
                    if node.get("machineAgentPresent", False):
                        numberNodesWithMachineAgentInstalled += 1

                        # Check if an App Agent is also present to count agents installed alongside each other.
                        # Assuming 'appAgentPresent' is always a key if 'node' is well-formed.
                        if node["appAgentPresent"]:
                            numberMachineAgentsInstalledAlongsideAppAgents += 1

                        # Safely retrieve the machine agent version string from the nested metadata.
                        # Using .get() with default empty dictionaries to prevent KeyError if intermediate keys are missing.
                        machine_agent_version_str = node.get('metadata', {}).get('applicationComponentNode', {}).get('machineAgentVersion')

                        if machine_agent_version_str: # Only proceed if the version string is not None or empty
                            # Update nodeVersionMap with the retrieved version string
                            if machine_agent_version_str in nodeVersionMap:
                                nodeVersionMap[machine_agent_version_str] += 1
                            else:
                                nodeVersionMap[machine_agent_version_str] = 1

                            logging.info(f'application {application["name"]} node data: {node}')

                            # Calculate version age
                            match = semanticVersionRegex.search(machine_agent_version_str)
                            if match:
                                version = match[0].split(".")
                                majorVersion = int(version[0])
                                minorVersion = int(version[1])

                                logging.info(f'major version: {majorVersion}, minor version: {minorVersion}')

                                hostInfo["machineAgentVersions"].add((majorVersion, minorVersion))
                                application["machineAgentVersions"].append(f"{version[0]}.{version[1]}")

                                if majorVersion == 4:  # Agents with version 4 and below will always fail.
                                    node["machineAgentAge"] = 3
                                else:
                                    years = currYear - majorVersion
                                    if minorVersion < currMonth:
                                        years += 1
                                    node["machineAgentAge"] = years

                                    if years <= 2:
                                        numberMachineAgentsLessThan2YearsOld += 1
                                    if years == 1:
                                        numberMachineAgentsLessThan1YearOld += 1
                            else:
                                logging.warning(f"Could not parse machine agent version from '{machine_agent_version_str}' for node {node.get('name', 'unknown')}. Skipping age calculation.")
                                node["machineAgentAge"] = None # Indicate that age couldn't be determined
                        else:
                            logging.warning(f"Machine agent present for node {node.get('name', 'unknown')}, but 'machineAgentVersion' is None or empty in metadata. Skipping version-dependent analysis.")
                            node["machineAgentAge"] = None # Indicate that age couldn't be determined

                        # Determine application load - this uses 'machineAgentAvailability' which is at the top level of 'node'
                        if node.get("machineAgentAvailability", 0) != 0:
                            numberMachineAgentsReportingData += 1
                    else:
                        logging.info(f"No machine agent present for node {node.get('name', 'unknown')}. Skipping machine agent specific analysis for this node.")
                        # The original code had a 'continue' here to skip the rest of the loop for this node
                        # if 'machineAgentPresent' was False. We maintain that behavior.
                        continue # Skip to the next node if no machine agent is present.

                # In the case of multiple versions, will return the largest common agent count regardless of version.
                try:
                    # Only attempt to find the max if nodeVersionMap is not empty
                    if nodeVersionMap:
                        numberMachineAgentsRunningSameVersion = nodeVersionMap[max(nodeVersionMap, key=nodeVersionMap.get)]
                    else:
                        numberMachineAgentsRunningSameVersion = 0 # No machine agents, so 0 running same version
                except ValueError:
                    # This ValueError typically happens if max() is called on an empty sequence,
                    # which is now explicitly handled by 'if nodeVersionMap:'. This block might be redundant but harmless.
                    logging.debug(
                        f'{hostInfo["controller"].host} - No machine agents returned for application {application["name"]}, unable to parse agent versions.'
                    )
                    numberMachineAgentsRunningSameVersion = 0 # Ensure it's initialized even on error

                # Calculate percentage of compliant nodes.
                # Default all to bronze. Will be modified in call to 'applyThresholds'.
                if numberNodesWithMachineAgentInstalled != 0.0:
                    analysisDataEvaluatedMetrics["percentAgentsLessThan1YearOld"] = (
                            numberMachineAgentsLessThan1YearOld / numberNodesWithMachineAgentInstalled * 100
                    )
                    analysisDataEvaluatedMetrics["percentAgentsLessThan2YearsOld"] = (
                            numberMachineAgentsLessThan2YearsOld / numberNodesWithMachineAgentInstalled * 100
                    )
                    analysisDataEvaluatedMetrics["percentAgentsReportingData"] = (
                            numberMachineAgentsReportingData / numberNodesWithMachineAgentInstalled * 100
                    )
                    analysisDataEvaluatedMetrics["percentAgentsRunningSameVersion"] = (
                            numberMachineAgentsRunningSameVersion / numberNodesWithMachineAgentInstalled * 100
                    )
                    analysisDataEvaluatedMetrics["percentAgentsInstalledAlongsideAppAgents"] = (
                            numberMachineAgentsInstalledAlongsideAppAgents / numberNodesWithMachineAgentInstalled * 100
                    )
                else:
                    analysisDataEvaluatedMetrics["percentAgentsLessThan1YearOld"] = 0
                    analysisDataEvaluatedMetrics["percentAgentsLessThan2YearsOld"] = 0
                    analysisDataEvaluatedMetrics["percentAgentsReportingData"] = 0
                    analysisDataEvaluatedMetrics["percentAgentsRunningSameVersion"] = 0
                    analysisDataEvaluatedMetrics["percentAgentsInstalledAlongsideAppAgents"] = 0

                analysisDataRawMetrics["numberOfNodes"] = len(application["nodes"])
                analysisDataRawMetrics["numberNodesWithMachineAgentInstalled"] = numberNodesWithMachineAgentInstalled
                analysisDataRawMetrics["numberMachineAgentsInstalledAlongsideAppAgents"] = numberMachineAgentsInstalledAlongsideAppAgents
                analysisDataRawMetrics["numberMachineAgentsLessThan1YearOld"] = numberMachineAgentsLessThan1YearOld
                analysisDataRawMetrics["numberMachineAgentsLessThan2YearsOld"] = numberMachineAgentsLessThan2YearsOld

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)