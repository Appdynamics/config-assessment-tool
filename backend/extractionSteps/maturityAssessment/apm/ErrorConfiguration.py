import logging
from collections import OrderedDict

from api.appd.AppDService import AppDService
from extractionSteps.JobStepBase import JobStepBase
from util.asyncio_utils import AsyncioUtils
from util.stdlib_utils import substringBetween


class ErrorConfiguration(JobStepBase):
    def __init__(self):
        super().__init__("apm")

    async def extract(self, controllerData):
        """
        Extract error configuration details.
        1. Makes one API call per application to get BT Errors Per Minute.
        2. Is dependent on BusinessTransactions report, in which businessTransactionCallsPerMinute is gathered.
        3. Is dependent on Overhead report, in which applicationConfiguration is gathered.
        """
        jobStepName = type(self).__name__

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Extracting {jobStepName}')
            controller: AppDService = hostInfo["controller"]

            # Gather necessary metrics.
            getBusinessTransactionErrorsPerMinuteFutures = []
            for application in hostInfo[self.componentType].values():
                getBusinessTransactionErrorsPerMinuteFutures.append(
                    controller.getMetricData(
                        applicationID=application["id"],
                        metric_path="Business Transaction Performance|Business Transactions|*|*|Errors per Minute",
                        rollup=True,
                        time_range_type="BEFORE_NOW",
                        duration_in_mins="1440",
                    )
                )

            businessTransactionErrorsPerMinute = await AsyncioUtils.gatherWithConcurrency(*getBusinessTransactionErrorsPerMinuteFutures)

            for idx, applicationName in enumerate(hostInfo[self.componentType]):
                application = hostInfo[self.componentType][applicationName]
                application["businessTransactionErrorsPerMinute"] = businessTransactionErrorsPerMinute[idx].data

    def analyze(self, controllerData, thresholds):
        """
        Analysis of error configuration details.
        1. Determines number of custom error detection rules.
        2. Determines success (non-error) rate of least successful transaction.
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

                # successPercentageOfWorstTransaction
                # Create BT calls per minute lookup table
                businessTransactionCallsPerMinuteMap = {}
                for businessTransaction in application["businessTransactionCallsPerMinute"]:
                    for metricValue in businessTransaction["metricValues"]:
                        btName = substringBetween(
                            businessTransaction["metricPath"],
                            left="Business Transaction Performance|Business Transactions|",
                            right="|Calls per Minute",
                        )
                        businessTransactionCallsPerMinuteMap[btName] = metricValue["sum"]
                # Iterate BT errors per minute to find worst performing BT
                highestErrorPercentageOfAnyBusinessTransaction = 0
                for businessTransaction in application["businessTransactionErrorsPerMinute"]:
                    for metricValue in businessTransaction["metricValues"]:
                        btName = substringBetween(
                            businessTransaction["metricPath"],
                            left="Business Transaction Performance|Business Transactions|",
                            right="|Errors per Minute",
                        )
                        try:
                            if businessTransactionCallsPerMinuteMap[btName] != 0:
                                errorRate = metricValue["sum"] / businessTransactionCallsPerMinuteMap[btName] * 100
                            else:
                                errorRate = 0
                            if errorRate > highestErrorPercentageOfAnyBusinessTransaction:
                                highestErrorPercentageOfAnyBusinessTransaction = errorRate
                        except KeyError:
                            logging.warning(f"{btName} did not return CPM data but did return EPM data")
                analysisDataEvaluatedMetrics["successPercentageOfWorstTransaction"] = 100 - highestErrorPercentageOfAnyBusinessTransaction

                # numberOfCustomRules
                errorConfigTypes = [
                    "errorConfig",  # java
                    "dotNetErrorConfig",
                    "phpErrorConfiguration",
                    "nodeJsErrorConfiguration",
                    "pythonErrorConfiguration",
                    "rubyErrorConfiguration",
                ]

                customerLoggerDefinitions = 0
                ignoreExceptions = 0
                ignoreLoggerMsgPatterns = 0
                ignoreLoggerNames = 0
                httpErrorReturnCodes = 0
                errorRedirectPages = 0
                for errorConfigType in errorConfigTypes:
                    errorConfig = application["applicationConfiguration"][errorConfigType]

                    customerLoggerDefinitions += (
                        len(errorConfig["customerLoggerDefinitions"]) if errorConfig["customerLoggerDefinitions"] is not None else 0
                    )
                    ignoreExceptions += len(errorConfig["ignoreExceptions"]) if errorConfig["ignoreExceptions"] is not None else 0
                    ignoreLoggerMsgPatterns += (
                        len(errorConfig["ignoreLoggerMsgPatterns"]) if errorConfig["ignoreLoggerMsgPatterns"] is not None else 0
                    )
                    ignoreLoggerNames += len(errorConfig["ignoreLoggerNames"]) if errorConfig["ignoreLoggerNames"] is not None else 0
                    httpErrorReturnCodes += len(errorConfig["httpErrorReturnCodes"]) if errorConfig["httpErrorReturnCodes"] is not None else 0
                    errorRedirectPages += len(errorConfig["errorRedirectPages"]) if errorConfig["errorRedirectPages"] is not None else 0

                analysisDataEvaluatedMetrics["numberOfCustomRules"] = (
                    customerLoggerDefinitions
                    + ignoreExceptions
                    + ignoreLoggerMsgPatterns
                    + ignoreLoggerNames
                    + httpErrorReturnCodes
                    + errorRedirectPages
                )

                analysisDataRawMetrics["numberOfCustomRules"] = analysisDataEvaluatedMetrics["numberOfCustomRules"]
                analysisDataRawMetrics["highestErrorPercentageOfAnyBusinessTransaction"] = highestErrorPercentageOfAnyBusinessTransaction
                analysisDataRawMetrics["customerLoggerDefinitions"] = customerLoggerDefinitions
                analysisDataRawMetrics["ignoreExceptions"] = ignoreExceptions
                analysisDataRawMetrics["ignoreLoggerMsgPatterns"] = ignoreLoggerMsgPatterns
                analysisDataRawMetrics["ignoreLoggerNames"] = ignoreLoggerNames
                analysisDataRawMetrics["httpErrorReturnCodes"] = httpErrorReturnCodes
                analysisDataRawMetrics["errorRedirectPages"] = errorRedirectPages

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)
