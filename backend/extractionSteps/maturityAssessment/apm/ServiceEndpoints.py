import logging
from collections import OrderedDict

from api.appd.AppDService import AppDService
from extractionSteps.JobStepBase import JobStepBase
from util.asyncio_utils import gatherWithConcurrency


class ServiceEndpoints(JobStepBase):
    def __init__(self):
        super().__init__("apm")

    async def extract(self, controllerData):
        """
        Extract service endpoint details.
        1. Makes one API call per application to get Service Endpoint Calls Per Minute.
        2. Makes one API call per application to get Service Endpoint Custom Match Rules.
        """
        jobStepName = type(self).__name__

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Extracting {jobStepName}')
            controller: AppDService = hostInfo["controller"]

            # Gather necessary metrics.
            getServiceEndpointCallsPerMinuteFutures = []
            getServiceEndpointMatchRulesFutures = []
            for application in hostInfo[self.componentType].values():
                getServiceEndpointCallsPerMinuteFutures.append(
                    controller.getMetricData(
                        applicationID=application["id"],
                        metric_path="Service Endpoints|*|*|Calls per Minute",
                        rollup=True,
                        time_range_type="BEFORE_NOW",
                        duration_in_mins="60",
                    )
                )
                getServiceEndpointMatchRulesFutures.append(controller.getServiceEndpointMatchRules(application["id"]))

            serviceEndpointCallsPerMinute = await gatherWithConcurrency(*getServiceEndpointCallsPerMinuteFutures)
            serviceEndpointMatchRules = await gatherWithConcurrency(*getServiceEndpointMatchRulesFutures)

            for idx, applicationName in enumerate(hostInfo[self.componentType]):
                application = hostInfo[self.componentType][applicationName]
                application["serviceEndpoints"] = serviceEndpointCallsPerMinute[idx].data
                application["serviceEndpointCustomMatchRules"] = serviceEndpointMatchRules[idx].data[0]
                application["serviceEndpointDefaultMatchRules"] = serviceEndpointMatchRules[idx].data[1]

    def analyze(self, controllerData, thresholds):
        """
        Analysis of node level details.
        1. Determines if the global Service Endpoint limit is hit.
        2. Determines
        """

        jobStepName = type(self).__name__

        # Get thresholds related to job
        jobStepThresholds = thresholds[self.componentType][jobStepName]

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Analyzing {jobStepName}')

            # Service endpoint limit is global
            totalServiceEndpoints = sum(len(application["serviceEndpoints"]) for application in hostInfo[self.componentType].values())
            serviceEndpointLimit = next(
                iter([property for property in hostInfo["configurations"] if property["name"] == "sep.ADD.registration.limit"]),
                None,
            )
            if serviceEndpointLimit is None:
                logging.warning(f'{hostInfo["controller"].host} - Unable to find property sep.ADD.registration.limit for controller.')
                serviceEndpointLimit = 0
            else:
                serviceEndpointLimit = int(serviceEndpointLimit["value"])
            serviceEndpointLimitHit = totalServiceEndpoints >= serviceEndpointLimit

            for application in hostInfo[self.componentType].values():
                # Root node of current application for current JobStep.
                analysisDataRoot = application[jobStepName] = OrderedDict()
                # This data goes into the 'JobStep - Metrics' xlsx sheet.
                analysisDataEvaluatedMetrics = analysisDataRoot["evaluated"] = OrderedDict()
                # This data goes into the 'JobStep - Raw' xlsx sheet.
                analysisDataRawMetrics = analysisDataRoot["raw"] = OrderedDict()

                numberOfServiceEndpoints = len(application["serviceEndpoints"])

                # numberOfCustomServiceEndpointRules
                numberOfCustomServiceEndpointRules = 0
                for tier in application["serviceEndpointCustomMatchRules"]:
                    if tier.error is None:
                        numberOfCustomServiceEndpointRules += len(tier.data)
                analysisDataEvaluatedMetrics["numberOfCustomServiceEndpointRules"] = numberOfCustomServiceEndpointRules

                # serviceEndpointLimitNotHit
                applicationContributingToSepLimit = numberOfServiceEndpoints > 0
                analysisDataEvaluatedMetrics["serviceEndpointLimitNotHit"] = not (applicationContributingToSepLimit and serviceEndpointLimitHit)

                # percentServiceEndpointsWithLoadOrDisabled
                numberOfServiceEndpointsWithLoad = 0
                for serviceEndpoint in application["serviceEndpoints"]:
                    for metricValue in serviceEndpoint["metricValues"]:
                        if metricValue["sum"] != 0:
                            numberOfServiceEndpointsWithLoad += 1
                serviceEndpointAutoDetectionEnabled = False
                for defaultRule in application["serviceEndpointDefaultMatchRules"]:
                    for ruleType in defaultRule.data:
                        if ruleType["enabled"]:
                            serviceEndpointAutoDetectionEnabled = True
                if serviceEndpointAutoDetectionEnabled:
                    if numberOfServiceEndpoints > 0:
                        analysisDataEvaluatedMetrics["percentServiceEndpointsWithLoadOrDisabled"] = (
                            numberOfServiceEndpointsWithLoad / numberOfServiceEndpoints * 100
                        )
                    else:
                        analysisDataEvaluatedMetrics["percentServiceEndpointsWithLoadOrDisabled"] = 0
                else:
                    analysisDataEvaluatedMetrics["percentServiceEndpointsWithLoadOrDisabled"] = 0

                analysisDataRawMetrics["numberOfServiceEndpoints"] = numberOfServiceEndpoints
                analysisDataRawMetrics["numberOfServiceEndpointsWithLoad"] = numberOfServiceEndpointsWithLoad
                analysisDataRawMetrics["numberOfCustomServiceEndpointRules"] = numberOfCustomServiceEndpointRules
                analysisDataRawMetrics["controllerWideServiceEndpointLimit"] = serviceEndpointLimit
                analysisDataRawMetrics["controllerWideTotalServiceEndpoints"] = totalServiceEndpoints
                analysisDataRawMetrics["applicationContributingToSepLimit"] = applicationContributingToSepLimit

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)
