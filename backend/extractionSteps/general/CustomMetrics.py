import logging
from collections import OrderedDict

from api.appd.AppDService import AppDService
from extractionSteps.JobStepBase import BSGJobStepBase
from util.asyncio_utils import gatherWithConcurrency


class CustomMetrics(BSGJobStepBase):
    def __init__(self):
        super().__init__("apm")

    async def extract(self, controllerData):
        """
        Extract application customization details.
        1. Makes one API call per application to get custom extensions.
        """
        jobStepName = type(self).__name__

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Extracting details for {jobStepName}')
            controller: AppDService = hostInfo["controller"]

            getTiersFutures = []
            for application in hostInfo[self.componentType].values():
                getTiersFutures.append(controller.getTiers(application["id"]))
            allTiers = await gatherWithConcurrency(*getTiersFutures)

            allCustomMetrics = []
            for application, tiers in zip(hostInfo[self.componentType].values(), allTiers):
                getCustomMetricsFutures = []
                for tier in tiers.data:
                    getCustomMetricsFutures.append(
                        controller.getCustomMetrics(
                            applicationID=application["id"],
                            tierName=tier["name"],
                        )
                    )
                customMetrics = await gatherWithConcurrency(*getCustomMetricsFutures)
                allCustomMetrics.append(customMetrics)

            for idx, applicationName in enumerate(hostInfo[self.componentType]):
                application = hostInfo[self.componentType][applicationName]
                customMetrics = set()
                for tier in allCustomMetrics[idx]:
                    for customMetric in tier.data:
                        customMetrics.add(customMetric["name"])
                application["customMetrics"] = customMetrics

    def analyze(self, controllerData, thresholds):
        pass
