import logging
from collections import OrderedDict

from extractionSteps.JobStepBase import BSGJobStepBase


class OverallAssessment(BSGJobStepBase):
    def __init__(self):
        super().__init__("apm")

    async def extract(self, applicationInformation):
        pass

    def analyze(self, applicationInformation, thresholds):
        """
        Analysis of overall results to determine classification
        """

        jobStepName = type(self).__name__

        jobStepNames = [
            "AppAgents",
            "MachineAgents",
            "BusinessTransactions",
            "Backends",
            "Overhead",
            "ServiceEndpoints",
            "ErrorConfiguration",
            "HealthRulesAndAlerting",
            "DataCollectors",
            "ApmDashboards",
        ]

        num_job_steps = len(jobStepNames)

        # Get thresholds related to job
        jobStepThresholds = thresholds[jobStepName]

        for host, hostInfo in applicationInformation.items():
            logging.info(f'{hostInfo["controller"].host} - Analyzing details for {jobStepName}')

            for application in hostInfo[self.componentType].values():

                # Root node of current application for current JobStep.
                analysisDataRoot = application[jobStepName] = OrderedDict()
                # This data goes into the 'JobStep - Metrics' xlsx sheet.
                analysisDataEvaluatedMetrics = analysisDataRoot["evaluated"] = OrderedDict()
                # This data goes into the 'JobStep - Raw' xlsx sheet.
                analysisDataRawMetrics = analysisDataRoot["raw"] = OrderedDict()

                num_silver = 0
                num_gold = 0
                num_platinum = 0

                for individualJobStepName in jobStepNames:
                    job_step_color = application[individualJobStepName]["computed"]
                    if job_step_color[0] == "silver":
                        num_silver = num_silver + 1
                    elif job_step_color[0] == "gold":
                        num_gold = num_gold + 1
                    elif job_step_color[0] == "platinum":
                        num_platinum = num_platinum + 1

                analysisDataEvaluatedMetrics["PercentageTotalPlatinum"] = num_platinum / num_job_steps * 100
                analysisDataEvaluatedMetrics["PercentageTotalGoldOrBetter"] = (num_platinum + num_gold) / num_job_steps * 100
                analysisDataEvaluatedMetrics["PercentageTotalSilverOrBetter"] = (num_platinum + num_gold + num_silver) / num_job_steps * 100

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)
