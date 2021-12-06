import logging
from collections import OrderedDict
from datetime import datetime

from extractionSteps.JobStepBase import BSGJobStepBase
from util.stdlib_utils import get_recursively


class ApmDashboards(BSGJobStepBase):
    def __init__(self):
        super().__init__("apm")

    async def extract(self, applicationInformation):
        """
        Extract Dashboard details.
        1. No API calls to make, simply associate dashboards with which applications they have widgets for.
        """
        jobStepName = type(self).__name__

        for host, hostInfo in applicationInformation.items():
            logging.info(f'{hostInfo["controller"].host} - Extracting details for {jobStepName}')

            for dashboard in hostInfo["exportedDashboards"]:
                dashboard["applicationNames"] = get_recursively(dashboard, "applicationName")
                dashboard["applicationIDs"] = get_recursively(dashboard, "applicationId")
                dashboard["adqlQueries"] = get_recursively(dashboard, "adqlQueryList")

            for idx, applicationName in enumerate(hostInfo[self.componentType]):
                application = hostInfo[self.componentType][applicationName]
                application["apmDashboards"] = []
                application["biqDashboards"] = []

                for dashboard in hostInfo["exportedDashboards"]:
                    if application["name"] in dashboard["applicationNames"]:
                        application["apmDashboards"].append(dashboard)
                    elif application["id"] in dashboard["applicationIDs"]:
                        application["apmDashboards"].append(dashboard)

                    if any(application["name"] in item for item in dashboard["adqlQueries"]):
                        application["biqDashboards"].append(dashboard)

    def analyze(self, applicationInformation, thresholds):
        """
        Analysis of node level details.
        1. Determines last modified date of dashboards per application.
        2. Determines number of dashboards per application.
        3. Determines number of dashboards with BiQ widgets per application.
        """

        jobStepName = type(self).__name__

        # Get thresholds related to job
        jobStepThresholds = thresholds[jobStepName]

        now = datetime.now()

        for host, hostInfo in applicationInformation.items():
            logging.info(f'{hostInfo["controller"].host} - Analyzing details for {jobStepName}')

            for application in hostInfo[self.componentType].values():
                # Root node of current application for current JobStep.
                analysisDataRoot = application[jobStepName] = OrderedDict()
                # This data goes into the 'JobStep - Metrics' xlsx sheet.
                analysisDataEvaluatedMetrics = analysisDataRoot["evaluated"] = OrderedDict()
                # This data goes into the 'JobStep - Raw' xlsx sheet.
                analysisDataRawMetrics = analysisDataRoot["raw"] = OrderedDict()

                # numberOfDashboards
                analysisDataEvaluatedMetrics["numberOfDashboards"] = len(application["apmDashboards"]) + len(application["biqDashboards"])

                # percentageOfDashboardsModifiedLast6Months
                numDashboardsModifiedLast6Months = 0
                for dashboard in application["apmDashboards"]:
                    modified = datetime.fromtimestamp(dashboard["modifiedOn"] / 1000.0)
                    num_months = (now.year - modified.year) * 12 + (now.month - modified.month)
                    if num_months <= 6:
                        numDashboardsModifiedLast6Months += 1
                for dashboard in application["biqDashboards"]:
                    modified = datetime.fromtimestamp(dashboard["modifiedOn"] / 1000.0)
                    num_months = (now.year - modified.year) * 12 + (now.month - modified.month)
                    if num_months <= 6:
                        numDashboardsModifiedLast6Months += 1
                if len(application["apmDashboards"]) + len(application["biqDashboards"]) == 0:
                    analysisDataEvaluatedMetrics["percentageOfDashboardsModifiedLast6Months"] = 0
                else:
                    analysisDataEvaluatedMetrics["percentageOfDashboardsModifiedLast6Months"] = (
                        numDashboardsModifiedLast6Months / (len(application["apmDashboards"]) + len(application["biqDashboards"])) * 100
                    )

                # numberOfDashboardsUsingBiQ
                analysisDataEvaluatedMetrics["numberOfDashboardsUsingBiQ"] = len(application["biqDashboards"])

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)
