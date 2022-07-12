import json
import logging
from collections import OrderedDict

from api.appd.AppDService import AppDService
from deepdiff import DeepDiff
from extractionSteps.JobStepBase import JobStepBase
from util.asyncio_utils import AsyncioUtils


class HealthRulesAndAlertingMRUM(JobStepBase):
    def __init__(self):
        super().__init__("mrum")

    async def extract(self, controllerData):
        """
        Extract health rule and alerting configuration details.
        1. Makes one API call per application to get Health Rules.
        2. Makes one API call per application to get Event Counts (health rule violations).
        3. Makes one API call per application to get Policies.
        """
        jobStepName = type(self).__name__

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Extracting {jobStepName}')
            controller: AppDService = hostInfo["controller"]

            # Gather necessary metrics.
            getHealthRulesFutures = []
            getEventCountsFutures = []
            getPoliciesFutures = []
            for application in hostInfo[self.componentType].values():
                getEventCountsFutures.append(
                    controller.getEventCountsLastDay(
                        applicationID=application["applicationId"],
                        entityType="APPLICATION",
                        entityID=application["applicationId"],
                    )
                )
                getHealthRulesFutures.append(controller.getHealthRules(application["applicationId"]))
                getPoliciesFutures.append(controller.getPolicies(application["applicationId"]))

            eventCounts = await AsyncioUtils.gatherWithConcurrency(*getEventCountsFutures)
            healthRules = await AsyncioUtils.gatherWithConcurrency(*getHealthRulesFutures)
            policies = await AsyncioUtils.gatherWithConcurrency(*getPoliciesFutures)

            for idx, applicationName in enumerate(hostInfo[self.componentType]):
                application = hostInfo[self.componentType][applicationName]

                application["eventCounts"] = eventCounts[idx].data
                application["policies"] = policies[idx].data

                trimmedHrs = [healthRule for healthRule in healthRules[idx].data if healthRule.error is None]
                application["healthRules"] = {
                    healthRuleList.data["name"]: healthRuleList.data for healthRuleList in trimmedHrs if healthRuleList.error is None
                }

    def analyze(self, controllerData, thresholds):
        """
        Analysis of error configuration details.
        1. Determines number of Health Rule violations in the last 24 hours.
        2. Determines number of Default Health Rules modified.
        3. Determines number of Actions currently bound to enabled policies.
        4. Determines number of Custom Health Rules.
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

                # numberOfHealthRuleViolationsLast24Hours
                policyEventCounts = application["eventCounts"]["policyViolationEventCounts"]["totalPolicyViolations"]
                analysisDataEvaluatedMetrics["numberOfHealthRuleViolationsLast24Hours"] = policyEventCounts["warning"] + policyEventCounts["critical"]

                # numberOfActionsBoundToEnabledPolicies
                actionsInEnabledPolicies = set()
                for policy in application["policies"]:
                    if policy["enabled"]:
                        if "actions" in policy:
                            for action in policy["actions"]:
                                actionsInEnabledPolicies.add(action["actionName"])
                        else:
                            logging.warning(f"Policy {policy['name']} is enabled but has no actions bound to it.")
                analysisDataEvaluatedMetrics["numberOfActionsBoundToEnabledPolicies"] = len(actionsInEnabledPolicies)

                # numberOfCustomHealthRules
                analysisDataEvaluatedMetrics["numberOfCustomHealthRules"] = len(application["healthRules"])

                analysisDataRawMetrics["totalWarningPolicyViolationsLast24Hours"] = policyEventCounts["warning"]
                analysisDataRawMetrics["totalCriticalPolicyViolationsLast24Hours"] = policyEventCounts["critical"]
                analysisDataRawMetrics["numberOfHealthRules"] = len(application["healthRules"])
                analysisDataRawMetrics["numberOfPolicies"] = len(application["policies"])

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)
