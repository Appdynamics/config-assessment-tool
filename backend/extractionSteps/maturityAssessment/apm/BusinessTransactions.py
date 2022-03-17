import logging
from collections import OrderedDict

from api.appd.AppDService import AppDService
from extractionSteps.JobStepBase import JobStepBase
from util.asyncio_utils import AsyncioUtils


class BusinessTransactions(JobStepBase):
    def __init__(self):
        super().__init__("apm")

    async def extract(self, controllerData):
        """
        Extract business transaction details.
        1. Makes one API call per application to get BT Calls Per Minute.
        2. Makes one API call per application to get BT Level Configuration (BT Lockdown enabled).
        3. Makes one API call per application to get BT Match Rules.
        """
        jobStepName = type(self).__name__

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Extracting {jobStepName}')
            controller: AppDService = hostInfo["controller"]

            # Gather necessary metrics.
            getBTCallsPerMinuteFutures = []
            getAppLevelBtConfigFutures = []
            getBtMatchRulesFutures = []
            for application in hostInfo[self.componentType].values():
                getBTCallsPerMinuteFutures.append(
                    controller.getMetricData(
                        applicationID=application["id"],
                        metric_path="Business Transaction Performance|Business Transactions|*|*|Calls per Minute",
                        rollup=True,
                        time_range_type="BEFORE_NOW",
                        duration_in_mins="1440",
                    )
                )
                getAppLevelBtConfigFutures.append(controller.getAppLevelBTConfig(application["id"]))
                getBtMatchRulesFutures.append(controller.getBtMatchRules(application["id"]))

            btCallsPerMinute = await AsyncioUtils.gatherWithConcurrency(*getBTCallsPerMinuteFutures)
            appLevelBtConfig = await AsyncioUtils.gatherWithConcurrency(*getAppLevelBtConfigFutures)
            btMatchRules = await AsyncioUtils.gatherWithConcurrency(*getBtMatchRulesFutures)

            for idx, applicationName in enumerate(hostInfo[self.componentType]):
                application = hostInfo[self.componentType][applicationName]
                application["businessTransactionCallsPerMinute"] = btCallsPerMinute[idx].data
                application["appLevelBtConfig"] = appLevelBtConfig[idx].data
                application["btMatchRules"] = btMatchRules[idx].data

    def analyze(self, controllerData, thresholds):
        """
        Analysis of node level details.
        1. Determines if BT limit is hit.
        2. Determines percentage for BTs with load.
        3. Determines if BT Lockdown is enabled.
        4. Determines number of custom BT match rules.
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

                # TODO: at least 1 business Transaction

                # btLimitNotHit
                numberOfBusinessTransactions = len(application["businessTransactionCallsPerMinute"])
                analysisDataEvaluatedMetrics["numberOfBTs"] = numberOfBusinessTransactions

                # percentBTsWithLoad
                businessTransactionsWithLoad = 0
                for businessTransaction in application["businessTransactionCallsPerMinute"]:
                    for metricValue in businessTransaction["metricValues"]:
                        if metricValue["sum"] != 0:
                            businessTransactionsWithLoad += 1
                if numberOfBusinessTransactions != 0:
                    analysisDataEvaluatedMetrics["percentBTsWithLoad"] = (businessTransactionsWithLoad / numberOfBusinessTransactions) * 100
                else:
                    analysisDataEvaluatedMetrics["percentBTsWithLoad"] = 0

                # isBtLockDownEnabled
                if "isBtLockDownEnabled" in application["appLevelBtConfig"]:
                    analysisDataEvaluatedMetrics["btLockdownEnabled"] = application["appLevelBtConfig"]["isBtLockDownEnabled"]
                else:
                    analysisDataEvaluatedMetrics["btLockdownEnabled"] = False

                # numberCustomMatchRules
                defaultBtNameList = [
                    "Web Server Auto Discovery Rule",
                    "Jersey 2.x Servlet",
                    "Node.js Auto Discovery Rule",
                    "Python Static Content Filter",
                    "Python Auto Discovery Rule",
                    "Weblogic JAX WS Webservice Servlet",
                    "JBoss 6.x web-services Servlet",
                    ".NET Auto Discovery Rule",
                    "Jersey Servlet",
                    "ASP.NET WebService Script Handler",
                    "JAX WS RI Dispatcher Servlet",
                    "Websphere web-services Servlet",
                    "Spring WS - dispatching of Web service messages",
                    "Weblogic JAX RPC Servlets",
                    "Spring WS - Base servlet for Spring's web framework",
                    "Php Auto Discovery Rule",
                    "Websphere web-services axis Servlet",
                    "Weblogic JAX WS Servlet",
                    "ASP.NET MVC5 Resource Handler",
                    "Java Auto Discovery Rule",
                    "XFire web-services servlet",
                    "CometD Annotation Servlet",
                    "CometD Servlet",
                    "JBoss web-services servlet",
                    "Apache Axis Servlet",
                    "Struts Action Servlet",
                    "Apache Axis2 Servlet",
                    "ASP.NET WebService Session Handler",
                    "NodeJS Static Content Filter",
                    "ASP.NET WCF Activation Handler",
                    "Apache Axis2 Admin Servlet",
                    "Quartz",
                ]
                numberOfCustomMatchRules = 0
                if "ruleScopeSummaryMappings" in application["btMatchRules"]:
                    for rule in application["btMatchRules"]["ruleScopeSummaryMappings"]:
                        if rule["rule"]["summary"]["name"] not in defaultBtNameList and rule["rule"]["enabled"]:
                            numberOfCustomMatchRules += 1
                analysisDataEvaluatedMetrics["numberCustomMatchRules"] = numberOfCustomMatchRules

                analysisDataRawMetrics["numberOfBTs"] = numberOfBusinessTransactions
                analysisDataRawMetrics["businessTransactionsWithLoad"] = businessTransactionsWithLoad
                analysisDataRawMetrics["btLockdownEnabled"] = analysisDataEvaluatedMetrics["btLockdownEnabled"]
                analysisDataRawMetrics["numberCustomMatchRules"] = numberOfCustomMatchRules

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)
