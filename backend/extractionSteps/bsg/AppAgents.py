import logging
import re
import time
from collections import OrderedDict

from api.appd.AppDService import AppDService
from extractionSteps.BSGJobStepBase import BSGJobStepBase
from util.asyncio_utils import gatherWithConcurrency
from util.stdlib_utils import substringBetween


class AppAgents(BSGJobStepBase):
    def __init__(self):
        super().__init__("apm")

    async def extract(self, controllerData):
        """
        Extract node level details.
        1. Makes one API call per application to get Node Metadata.
        2. Makes one API call per application to get Node App Agent Availability in the last 24 hours.
        3. Makes one API call per application to get Node Requests Exceeding Limit in the last 24 hours.
        """
        jobStepName = type(self).__name__

        nodeIdToAppAgentAvailabilityMap = {}
        nodeIdToMetricLimitMap = {}
        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Extracting details for {jobStepName}')
            controller: AppDService = hostInfo["controller"]

            # Gather necessary metrics.
            getNodesFutures = []
            appAgentAvailabilityFutures = []
            nodeMetricsUploadRequestsExceedingLimitFutures = []
            for application in hostInfo[self.componentType].values():
                getNodesFutures.append(controller.getNodes(application["id"]))
                appAgentAvailabilityFutures.append(
                    controller.getMetricData(
                        applicationID=application["id"],
                        metric_path="Application Infrastructure Performance|*|Individual Nodes|*|Agent|App|Availability",
                        rollup=True,
                        time_range_type="BEFORE_NOW",
                        duration_in_mins="60",
                    )
                )
                nodeMetricsUploadRequestsExceedingLimitFutures.append(
                    controller.getMetricData(
                        applicationID=application["id"],
                        metric_path="Application Infrastructure Performance|*|Individual Nodes|*|Agent|Metric Upload|Requests Exceeding Limit",
                        rollup=True,
                        time_range_type="BEFORE_NOW",
                        duration_in_mins="60",
                    )
                )
            nodes = await gatherWithConcurrency(*getNodesFutures)
            appAgentAvailability = await gatherWithConcurrency(*appAgentAvailabilityFutures)
            nodeMetricsUploadRequestsExceedingLimit = await gatherWithConcurrency(*nodeMetricsUploadRequestsExceedingLimitFutures)

            # Create a dictionary of Node -> Calls Per Minute for fast lookup
            for rolledUpMetrics in appAgentAvailability:
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
                        appAgentAvailability = nodeMetric["metricValues"][0]["sum"]
                    except IndexError:
                        tierName = ""
                        nodeName = ""
                        appAgentAvailability = 0
                    nodeIdToAppAgentAvailabilityMap[tierName + "|" + nodeName] = appAgentAvailability

            # Create a dictionary of Node -> Metrics Upload Requests Exceeding Limit for fast lookup
            for rolledUpMetrics in nodeMetricsUploadRequestsExceedingLimit:
                if rolledUpMetrics.error is not None:  # call to gather metrics failed for some reason (most likely 504)
                    continue
                for nodeMetric in rolledUpMetrics.data:
                    try:
                        # e.g. 'Application Infrastructure Performance|foo|Individual Nodes|bar|Agent|Metric Upload|Requests Exceeding Limit'
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
                        nodeMetricsUploadedExceedingLimitCount = nodeMetric["metricValues"][0]["sum"]
                    except IndexError:
                        tierName = ""
                        nodeName = ""
                        nodeMetricsUploadedExceedingLimitCount = 0
                    nodeIdToMetricLimitMap[tierName + "|" + nodeName] = nodeMetricsUploadedExceedingLimitCount

            # Append node level information to overall host info
            for idx, application in enumerate(hostInfo[self.componentType]):
                hostInfo[self.componentType][application]["nodes"] = nodes[idx].data
                for node in nodes[idx].data:
                    try:
                        node["appAgentAvailabilityLast24Hours"] = nodeIdToAppAgentAvailabilityMap[node["tierName"] + "|" + node["name"]]
                    except KeyError:
                        node["appAgentAvailabilityLast24Hours"] = 0
                        logging.debug(
                            f'{hostInfo["controller"].host} - Node: {node["tierName"]}|{node["name"]} returned no metric data for Agent Availability.'
                        )

                    try:
                        node["nodeMetricsUploadRequestsExceedingLimit"] = nodeIdToMetricLimitMap[node["tierName"] + "|" + node["name"]]
                    except KeyError:
                        node["nodeMetricsUploadRequestsExceedingLimit"] = 0
                        logging.debug(
                            f'{hostInfo["controller"].host} - Node: {node["tierName"]}|{node["name"]} returned no metric data for Metrics Upload Requests Exceeding Limit.'
                        )

    def analyze(self, controllerData, thresholds):
        """
        Analysis of node level details.
        1. Determines App Agent age from semantic versioning. (Version 4.X and under will always fail).
        2. Determines number of App Agents reporting data.
        3. Determines number of App Agents running same version. In the case of multiple versions, will return the largest common agent count regardless of version.
        4. Determines if any node in the application is hitting the metric limit.
        """

        # Used to determine agent age from semantic versioning of agents
        currYearAndMonth = [int(x) for x in time.strftime("%Y,%m").split(",")]
        currYear = int(str(currYearAndMonth[0])[-2:])
        currMonth = currYearAndMonth[1]

        semanticVersionRegex = re.compile("[0-9]+\\.[0-9]+\\.")

        jobStepName = type(self).__name__

        # Get thresholds related to job
        jobStepThresholds = thresholds[jobStepName]

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Analyzing details for {jobStepName}')

            hostInfo["appAgentVersions"] = set()

            for application in hostInfo[self.componentType].values():
                # Root node of current application for current JobStep.
                analysisDataRoot = application[jobStepName] = OrderedDict()
                # This data goes into the 'JobStep - Metrics' xlsx sheet.
                analysisDataEvaluatedMetrics = analysisDataRoot["evaluated"] = OrderedDict()
                # This data goes into the 'JobStep - Raw' xlsx sheet.
                analysisDataRawMetrics = analysisDataRoot["raw"] = OrderedDict()

                numberNodesWithAppAgentInstalled = 0.0
                numberAppAgentsLessThan1YearOld = 0
                numberAppAgentsLessThan2YearsOld = 0
                numberAppAgentsReportingData = 0
                numberAppAgentsRunningSameVersion = 0
                analysisDataEvaluatedMetrics["metricLimitNotHit"] = True
                nodeVersionMap = {}

                application["appAgentVersions"] = []

                for node in application["nodes"]:
                    if node["appAgentVersion"] in nodeVersionMap:
                        nodeVersionMap[node["appAgentVersion"]] += 1
                    else:
                        nodeVersionMap[node["appAgentVersion"]] = 1

                    # Calculate version age
                    numberNodesWithAppAgentInstalled += 1
                    if "" == node["appAgentVersion"]:  # No agent installed
                        continue

                    version = semanticVersionRegex.search(node["appAgentVersion"])[0].split(".")  # e.g. 'Server Agent v21.6.1.2 GA ...'
                    majorVersion = int(version[0])
                    minorVersion = int(version[1])

                    hostInfo["appAgentVersions"].add((majorVersion, minorVersion, node["agentType"]))
                    application["appAgentVersions"].append(f"{node['agentType']}:{version[0]}.{version[1]}")

                    if majorVersion == 4:  # Agents with version 4 and below will always fail.
                        node["appAgentAge"] = 3
                    else:
                        years = currYear - majorVersion
                        if minorVersion < currMonth:
                            years += 1
                        node["appAgentAge"] = years

                        if years <= 2:
                            numberAppAgentsLessThan2YearsOld += 1
                        if years == 1:
                            numberAppAgentsLessThan1YearOld += 1

                    # Determine application load
                    if node["appAgentAvailabilityLast24Hours"] != 0:
                        numberAppAgentsReportingData += 1

                    if node["nodeMetricsUploadRequestsExceedingLimit"] != 0:
                        analysisDataEvaluatedMetrics["metricLimitNotHit"] = False

                # In the case of multiple versions, will return the largest common agent count regardless of version.
                try:
                    numberAppAgentsRunningSameVersion = nodeVersionMap[max(nodeVersionMap, key=nodeVersionMap.get)]
                except ValueError:
                    logging.debug(
                        f'{hostInfo["controller"].host} - No app agents returned for application {application["name"]}, unable to parse agent versions.'
                    )

                # Calculate percents of compliant nodes.
                if numberNodesWithAppAgentInstalled != 0.0:
                    analysisDataEvaluatedMetrics["percentAgentsLessThan1YearOld"] = (
                        numberAppAgentsLessThan1YearOld / numberNodesWithAppAgentInstalled * 100
                    )
                    analysisDataEvaluatedMetrics["percentAgentsLessThan2YearsOld"] = (
                        numberAppAgentsLessThan2YearsOld / numberNodesWithAppAgentInstalled * 100
                    )
                    analysisDataEvaluatedMetrics["percentAgentsReportingData"] = numberAppAgentsReportingData / numberNodesWithAppAgentInstalled * 100
                    analysisDataEvaluatedMetrics["percentAgentsRunningSameVersion"] = (
                        numberAppAgentsRunningSameVersion / numberNodesWithAppAgentInstalled * 100
                    )
                else:
                    analysisDataEvaluatedMetrics["percentAgentsLessThan1YearOld"] = 0
                    analysisDataEvaluatedMetrics["percentAgentsLessThan2YearsOld"] = 0
                    analysisDataEvaluatedMetrics["percentAgentsReportingData"] = 0
                    analysisDataEvaluatedMetrics["percentAgentsRunningSameVersion"] = 0

                analysisDataRawMetrics["numberOfNodes"] = len(application["nodes"])
                analysisDataRawMetrics["numberNodesWithAppAgentInstalled"] = numberNodesWithAppAgentInstalled
                analysisDataRawMetrics["numberAppAgentsLessThan1YearOld"] = numberAppAgentsLessThan1YearOld
                analysisDataRawMetrics["numberAppAgentsLessThan2YearsOld"] = numberAppAgentsLessThan2YearsOld
                analysisDataRawMetrics["numberOfAgentsReportingData"] = numberAppAgentsReportingData

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)
