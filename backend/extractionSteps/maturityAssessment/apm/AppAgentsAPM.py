import logging
import re
import time
from collections import OrderedDict

from backend.api.appd.AppDService import AppDService
from backend.extractionSteps.JobStepBase import JobStepBase
from util.asyncio_utils import AsyncioUtils
from util.stdlib_utils import substringBetween


logger = logging.getLogger(__name__.split('.')[-1])


class AppAgentsAPM(JobStepBase):
    def __init__(self):
        super().__init__("apm")

    async def extract(self, controllerData):
        """
        Extract node level details.
        1. Makes one API call per application to get Node Metadata.
        2. Makes one API call per application to get Node App Agent Availability.
        3. Makes one API call per application to get Node Requests Exceeding Limit.
        """
        jobStepName = type(self).__name__

        nodeIdToAppAgentAvailabilityMap = {}
        nodeIdToMetricLimitMap = {}
        for host, hostInfo in controllerData.items():
            logger.info(f'{hostInfo["controller"].host} - Extracting {jobStepName}')
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
                        duration_in_mins=controller.timeRangeMins,
                    )
                )
                nodeMetricsUploadRequestsExceedingLimitFutures.append(
                    controller.getMetricData(
                        applicationID=application["id"],
                        metric_path="Application Infrastructure Performance|*|Individual Nodes|*|Agent|Metric Upload|Requests Exceeding Limit",
                        rollup=True,
                        time_range_type="BEFORE_NOW",
                        duration_in_mins=controller.timeRangeMins,
                    )
                )
            nodes = await AsyncioUtils.gatherWithConcurrency(*getNodesFutures)
            appAgentAvailability = await AsyncioUtils.gatherWithConcurrency(*appAgentAvailabilityFutures)
            nodeMetricsUploadRequestsExceedingLimit = await AsyncioUtils.gatherWithConcurrency(*nodeMetricsUploadRequestsExceedingLimitFutures)

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

            nodeMetadataFutures = []
            for nodesList, application in zip(nodes, (hostInfo[self.componentType].values())):
                nodeIds = [node["id"] for node in nodesList.data]
                nodeMetadataFutures.append(controller.getAppAgentMetadata(application["id"], nodeIds))
            nodeMetadata = await AsyncioUtils.gatherWithConcurrency(*nodeMetadataFutures)

            # Append node level information to overall host info
            hostInfo["nodeIdAppAgentAvailabilityMap"] = {}
            hostInfo["nodeIdMetaInfoMap"] = {}
            for idx, application in enumerate(hostInfo[self.componentType]):
                hostInfo[self.componentType][application]["nodes"] = nodes[idx].data
                for node, metadata in zip(nodes[idx].data, nodeMetadata[idx].data):
                    node["metadata"] = metadata
                    try:
                        node["appAgentAvailability"] = nodeIdToAppAgentAvailabilityMap[node["tierName"] + "|" + node["name"]]
                    except (KeyError, TypeError):
                        node["appAgentAvailability"] = 0
                        logger.debug(
                            f'{hostInfo["controller"].host} - Node: {node["tierName"]}|{node["name"]} returned no metric data for Agent Availability.'
                        )
                    hostInfo["nodeIdAppAgentAvailabilityMap"][node["id"]] = node["appAgentAvailability"] / controller.timeRangeMins * 100
                    hostInfo["nodeIdMetaInfoMap"][node["id"]] = node["metadata"]

                    try:
                        node["nodeMetricsUploadRequestsExceedingLimit"] = nodeIdToMetricLimitMap[node["tierName"] + "|" + node["name"]]
                    except (KeyError, TypeError):
                        node["nodeMetricsUploadRequestsExceedingLimit"] = 0
                        logger.debug(
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
        jobStepThresholds = thresholds[self.componentType][jobStepName]

        for host, hostInfo in controllerData.items():
            logger.info(f'{hostInfo["controller"].host} - Analyzing {jobStepName}')

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
                    # Support both APIs: new (explicit flag) and old (implicit via non-empty version string)
                    app_agent_present_flag = node.get("appAgentPresent", None)
                    app_agent_present = (
                        app_agent_present_flag is True
                        or (app_agent_present_flag is None and node.get("appAgentVersion", "") != "")
                    )

                    if app_agent_present:
                        version_str = node.get("appAgentVersion", "")
                        if version_str in nodeVersionMap:
                            nodeVersionMap[version_str] += 1
                        else:
                            nodeVersionMap[version_str] = 1

                        numberNodesWithAppAgentInstalled += 1

                        match = semanticVersionRegex.search(version_str)
                        if not match:
                            continue  # Cannot parse semantic version, skip aging logic

                        version = match[0].split(".")
                        majorVersion = int(version[0])
                        minorVersion = int(version[1])

                        hostInfo["appAgentVersions"].add((majorVersion, minorVersion, node.get("agentType")))
                        application["appAgentVersions"].append(f"{node.get('agentType')}:{version[0]}.{version[1]}")

                        if majorVersion == 4:
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

                        if node.get("appAgentAvailability", 0) != 0:
                            numberAppAgentsReportingData += 1

                        if node.get("nodeMetricsUploadRequestsExceedingLimit", 0) != 0:
                            analysisDataEvaluatedMetrics["metricLimitNotHit"] = False

                # In the case of multiple versions, will return the largest common agent count regardless of version.
                try:
                    numberAppAgentsRunningSameVersion = nodeVersionMap[max(nodeVersionMap, key=nodeVersionMap.get)]
                except ValueError:
                    logger.debug(
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
                analysisDataRawMetrics["numberOfTiers"] = len(application["tiers"])
                analysisDataRawMetrics["numberNodesWithAppAgentInstalled"] = numberNodesWithAppAgentInstalled
                analysisDataRawMetrics["numberAppAgentsLessThan1YearOld"] = numberAppAgentsLessThan1YearOld
                analysisDataRawMetrics["numberAppAgentsLessThan2YearsOld"] = numberAppAgentsLessThan2YearsOld
                analysisDataRawMetrics["numberOfAgentsReportingData"] = numberAppAgentsReportingData

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)
