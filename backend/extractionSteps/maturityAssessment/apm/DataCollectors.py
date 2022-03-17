import logging
from collections import OrderedDict

from api.appd.AppDService import AppDService
from extractionSteps.JobStepBase import JobStepBase
from util.asyncio_utils import AsyncioUtils


class DataCollectors(JobStepBase):
    def __init__(self):
        super().__init__("apm")

    async def extract(self, controllerData):
        """
        Extract node level details.
        1. Makes one API call per application to get Data Collectors.
        2. Makes one API call per Data Collector to get snapshots containing said Data Collector (max 1 result returned).
        """
        jobStepName = type(self).__name__

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Extracting {jobStepName}')
            controller: AppDService = hostInfo["controller"]

            # Gather necessary metrics.
            getDataCollectorsFutures = []

            for application in hostInfo[self.componentType].values():
                getDataCollectorsFutures.append(controller.getDataCollectorUsage(application["id"]))

            dataCollectors = await AsyncioUtils.gatherWithConcurrency(*getDataCollectorsFutures)

            for idx, applicationName in enumerate(hostInfo[self.componentType]):
                application = hostInfo[self.componentType][applicationName]
                application["dataCollectors"] = dataCollectors[idx].data

    def analyze(self, controllerData, thresholds):
        """
        Analysis of node level details.
        1. Determines number of Data Collector Fields.
        """

        jobStepName = type(self).__name__

        # Get thresholds related to job
        jobStepThresholds = thresholds[self.componentType][jobStepName]

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Analyzing {jobStepName}')

            for application in hostInfo[self.componentType].values():
                # Root node of current application for current JobStep.
                analysisDataRoot = application[jobStepName] = OrderedDict()
                # This data goes into the 'JobStep - Metrics' xlsx sheet.
                analysisDataEvaluatedMetrics = analysisDataRoot["evaluated"] = OrderedDict()
                # This data goes into the 'JobStep - Raw' xlsx sheet.
                analysisDataRawMetrics = analysisDataRoot["raw"] = OrderedDict()

                # numberOfDataCollectorFieldsConfigured
                analysisDataEvaluatedMetrics["numberOfDataCollectorFieldsConfigured"] = len(application["dataCollectors"]["allDataCollectors"])

                # numberOfDataCollectorFieldsCollectedInSnapshots
                analysisDataEvaluatedMetrics["numberOfDataCollectorFieldsCollectedInSnapshotsLast1Day"] = len(
                    application["dataCollectors"]["dataCollectorsPresentInSnapshots"]
                )

                # numberOfDataCollectorFieldsCollectedInAnalytics
                analysisDataEvaluatedMetrics["numberOfDataCollectorFieldsCollectedInAnalyticsLast1Day"] = len(
                    application["dataCollectors"]["dataCollectorsPresentInAnalytics"]
                )

                # biqEnabled
                biqEnabled = next(
                    iter([status["enabled"] for status in hostInfo["analyticsEnabledStatus"] if status["applicationId"] == application["id"]]),
                    False,
                )
                analysisDataEvaluatedMetrics["biqEnabled"] = biqEnabled

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)
