import logging
from collections import OrderedDict
from itertools import count

from api.appd.AppDService import AppDService
from extractionSteps.JobStepBase import JobStepBase
from util.asyncio_utils import AsyncioUtils


class NetworkRequests(JobStepBase):
    def __init__(self):
        super().__init__("brum")

    async def extract(self, controllerData):
        """
        Extract node level details.
        1.
        """
        jobStepName = type(self).__name__

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Extracting {jobStepName}')
            controller: AppDService = hostInfo["controller"]

            getEumPageListViewDataFutures = []
            getEumNetworkRequestListFutures = []
            getPagesAndFramesConfigFutures = []
            getAJAXConfigFutures = []
            getVirtualPagesConfigFutures = []
            for application in hostInfo[self.componentType].values():
                getEumPageListViewDataFutures.append(controller.getEumPageListViewData(application["id"]))
                getEumNetworkRequestListFutures.append(controller.getEumNetworkRequestList(application["id"]))
                getPagesAndFramesConfigFutures.append(controller.getPagesAndFramesConfig(application["id"]))
                getAJAXConfigFutures.append(controller.getAJAXConfig(application["id"]))
                getVirtualPagesConfigFutures.append(controller.getVirtualPagesConfig(application["id"]))

            eumPageListViewData = await AsyncioUtils.gatherWithConcurrency(*getEumPageListViewDataFutures)
            eumNetworkRequestList = await AsyncioUtils.gatherWithConcurrency(*getEumNetworkRequestListFutures)
            pagesAndFramesConfig = await AsyncioUtils.gatherWithConcurrency(*getPagesAndFramesConfigFutures)
            ajaxConfig = await AsyncioUtils.gatherWithConcurrency(*getAJAXConfigFutures)
            virtualPagesConfig = await AsyncioUtils.gatherWithConcurrency(*getVirtualPagesConfigFutures)

            for idx, application in enumerate(hostInfo[self.componentType]):
                hostInfo[self.componentType][application]["eumPageListViewData"] = eumPageListViewData[idx].data
                hostInfo[self.componentType][application]["eumNetworkRequestList"] = eumNetworkRequestList[idx].data
                hostInfo[self.componentType][application]["pagesAndFramesConfig"] = pagesAndFramesConfig[idx].data
                hostInfo[self.componentType][application]["ajaxConfig"] = ajaxConfig[idx].data
                hostInfo[self.componentType][application]["virtualPagesConfig"] = virtualPagesConfig[idx].data

    def analyze(self, controllerData, thresholds):
        """
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

                analysisDataEvaluatedMetrics["collectingDataPastOneDay"] = application["metrics"]["pageRequestsPerMin"]["sum"] > 0
                analysisDataRawMetrics["totalCallsPastOneDay"] = application["metrics"]["pageRequestsPerMin"]["sum"]

                analysisDataRawMetrics["totalNetworkRequests"] = application["eumNetworkRequestList"]["totalCount"]
                analysisDataRawMetrics["totalAJAXRequests"] = len(
                    [app for app in application["eumNetworkRequestList"]["data"] if app["type"] == "AJAX_REQUEST"]
                )
                analysisDataRawMetrics["totalBasePageRequests"] = len(
                    [app for app in application["eumNetworkRequestList"]["data"] if app["type"] == "BASE_PAGE"]
                )
                analysisDataRawMetrics["totalVirtualPageRequests"] = len(
                    [app for app in application["eumNetworkRequestList"]["data"] if app["type"] == "VIRTUAL_PAGE"]
                )
                analysisDataRawMetrics["totalIFrameRequests"] = len(
                    [app for app in application["eumNetworkRequestList"]["data"] if app["type"] == "IFRAME"]
                )

                analysisDataRawMetrics["pageIFrameLimit"] = application["eumPageListViewData"]["pageIFrameLimit"]
                analysisDataRawMetrics["ajaxLimit"] = application["eumPageListViewData"]["ajaxLimit"]

                analysisDataEvaluatedMetrics["networkRequestLimitNotHit"] = (
                    analysisDataRawMetrics["totalAJAXRequests"] < analysisDataRawMetrics["ajaxLimit"]
                    and (
                        analysisDataRawMetrics["totalBasePageRequests"]
                        + analysisDataRawMetrics["totalVirtualPageRequests"]
                        + analysisDataRawMetrics["totalIFrameRequests"]
                    )
                    < analysisDataRawMetrics["pageIFrameLimit"]
                )

                numberOfCustomPageIncludeRules = len(application["pagesAndFramesConfig"]["customNamingIncludeRules"])
                numberOfCustomPageExcludeRules = len(application["pagesAndFramesConfig"]["customNamingExcludeRules"])
                numberOfCustomAJAXIncludeRules = len(application["ajaxConfig"]["customNamingIncludeRules"])
                numberOfCustomAJAXExcludeRules = len(application["ajaxConfig"]["customNamingExcludeRules"])
                numberOfCustomVirtualIncludeRules = len(application["virtualPagesConfig"]["customNamingIncludeRules"])
                numberOfCustomVirtualExcludeRules = len(application["virtualPagesConfig"]["customNamingExcludeRules"])

                analysisDataEvaluatedMetrics["numberCustomMatchRules"] = (
                    numberOfCustomPageIncludeRules
                    + numberOfCustomPageExcludeRules
                    + numberOfCustomAJAXIncludeRules
                    + numberOfCustomAJAXExcludeRules
                    + numberOfCustomVirtualIncludeRules
                    + numberOfCustomVirtualExcludeRules
                )
                analysisDataRawMetrics["numberOfCustomPageIncludeRules"] = numberOfCustomPageIncludeRules
                analysisDataRawMetrics["numberOfCustomPageExcludeRules"] = numberOfCustomPageExcludeRules
                analysisDataRawMetrics["numberOfCustomAJAXIncludeRules"] = numberOfCustomAJAXIncludeRules
                analysisDataRawMetrics["numberOfCustomAJAXExcludeRules"] = numberOfCustomAJAXExcludeRules
                analysisDataRawMetrics["numberOfCustomVirtualIncludeRules"] = numberOfCustomVirtualIncludeRules
                analysisDataRawMetrics["numberOfCustomVirtualExcludeRules"] = numberOfCustomVirtualExcludeRules

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)
