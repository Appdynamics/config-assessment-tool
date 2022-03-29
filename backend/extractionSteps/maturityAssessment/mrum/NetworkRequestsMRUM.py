import logging
from collections import OrderedDict
from itertools import count

from api.appd.AppDService import AppDService
from extractionSteps.JobStepBase import JobStepBase
from util.asyncio_utils import AsyncioUtils


class NetworkRequestsMRUM(JobStepBase):
    def __init__(self):
        super().__init__("mrum")

    async def extract(self, controllerData):
        """
        Extract node level details.
        1.
        """
        jobStepName = type(self).__name__

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Extracting {jobStepName}')
            controller: AppDService = hostInfo["controller"]

            getMRUMNetworkRequestConfigFutures = []
            getNetworkRequestLimitFutures = []
            for application in hostInfo[self.componentType].values():
                getMRUMNetworkRequestConfigFutures.append(controller.getMRUMNetworkRequestConfig(application["applicationId"]))
                getNetworkRequestLimitFutures.append(controller.getNetworkRequestLimit(application["mobileAppId"]))

            mrumNetworkRequestConfigs = await AsyncioUtils.gatherWithConcurrency(*getMRUMNetworkRequestConfigFutures)
            networkRequestLimits = await AsyncioUtils.gatherWithConcurrency(*getNetworkRequestLimitFutures)

            for idx, applicationName in enumerate(hostInfo[self.componentType]):
                application = hostInfo[self.componentType][applicationName]

                application["eumPageListViewData"] = mrumNetworkRequestConfigs[idx].data
                application["networkRequestLimit"] = networkRequestLimits[idx].data

    def analyze(self, controllerData, thresholds):
        """k
        Analysis of node level details.
        1. Determines if Developer Mode is either enabled application wide or for any BT.
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

                analysisDataEvaluatedMetrics["collectingDataPastOneDay"] = application["metrics"]["networkRequestsPerMin"]["sum"] > 0
                analysisDataRawMetrics["totalCallsPastOneDay"] = application["metrics"]["networkRequestsPerMin"]["sum"]

                analysisDataRawMetrics["isExceeded"] = application["networkRequestLimit"]["isExceeded"]
                analysisDataRawMetrics["perEumAppLimit"] = application["networkRequestLimit"]["perEumAppLimit"]
                analysisDataRawMetrics["perMobileAppLimit"] = application["networkRequestLimit"]["perMobileAppLimit"]
                analysisDataRawMetrics["numberOfAddsForMobileApp"] = application["networkRequestLimit"]["numOfAddsForMobileApp"]
                analysisDataRawMetrics["numberOfAddsForEumApp"] = application["networkRequestLimit"]["numOfAddsForEumApp"]

                analysisDataEvaluatedMetrics["networkRequestLimitNotHit"] = not application["networkRequestLimit"]["isExceeded"]

                numberOfCustomIncludeRules = len(application["eumPageListViewData"]["customNamingIncludeRules"])
                numberOfCustomExcludeRules = len(application["eumPageListViewData"]["customNamingExcludeRules"])

                analysisDataEvaluatedMetrics["numberCustomMatchRules"] = numberOfCustomIncludeRules + numberOfCustomExcludeRules

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)
