import json
import logging
from collections import OrderedDict
from deepdiff import DeepDiff

from api.appd.AppDService import AppDService
from jobs.JobStepBase import JobStepBase
from util.asyncio_utils import gatherWithConcurrency


class HealthRulesAndAlerting(JobStepBase):
    def __init__(self):
        super().__init__("apm")

    async def extract(self, applicationInformation):
        """
        Extract health rule and alerting configuration details.
        1. Makes one API call per application to get Health Rules.
        2. Makes one API call per application to get Event Counts (health rule violations).
        3. Makes one API call per application to get Policies.
        """
        jobStepName = type(self).__name__

        for host, hostInfo in applicationInformation.items():
            logging.info(f'{hostInfo["controller"].host} - Extracting details for {jobStepName}')
            controller: AppDService = hostInfo["controller"]

            # Gather necessary metrics.
            getHealthRulesFutures = []
            getEventCountsFutures = []
            getPoliciesFutures = []
            for application in hostInfo[self.componentType].values():
                getEventCountsFutures.append(
                    controller.getEventCountsLastDay(
                        applicationID=application["id"],
                        entityType="APPLICATION",
                        entityID=application["id"],
                    )
                )
                getHealthRulesFutures.append(controller.getHealthRules(application["id"]))
                getPoliciesFutures.append(controller.getPolicies(application["id"]))

            eventCounts = await gatherWithConcurrency(*getEventCountsFutures)
            healthRules = await gatherWithConcurrency(*getHealthRulesFutures)
            policies = await gatherWithConcurrency(*getPoliciesFutures)

            for idx, applicationName in enumerate(hostInfo[self.componentType]):
                application = hostInfo[self.componentType][applicationName]

                application["eventCounts"] = eventCounts[idx].data
                application["policies"] = policies[idx].data

                trimmedHrs = [healthRule for healthRule in healthRules[idx].data if healthRule.error is None]
                application["healthRules"] = {
                    healthRuleList.data["name"]: healthRuleList.data for healthRuleList in trimmedHrs if healthRuleList.error is None
                }

    def analyze(self, applicationInformation, thresholds):
        """
        Analysis of error configuration details.
        1. Determines number of Health Rule violations in the last 24 hours.
        2. Determines number of Default Health Rules modified.
        3. Determines number of Actions currently bound to enabled policies.
        4. Determines number of Custom Health Rules.
        """

        jobStepName = type(self).__name__

        # Get thresholds related to job
        jobStepThresholds = thresholds[jobStepName]

        defaultHealthRules = json.loads(open("backend/resources/controllerDefaults/defaultHealthRules.json").read())
        for host, hostInfo in applicationInformation.items():
            logging.info(f'{hostInfo["controller"].host} - Analyzing details for {jobStepName}')

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

                # numberOfDefaultHealthRulesModified
                defaultHealthRulesModified = 0
                for hrName, heathRule in defaultHealthRules.items():
                    if hrName in application["healthRules"]:
                        del application["healthRules"][hrName]["id"]
                        healthRuleDiff = DeepDiff(
                            defaultHealthRules[hrName],
                            application["healthRules"][hrName],
                            ignore_order=True,
                        )
                        if healthRuleDiff != {}:
                            defaultHealthRulesModified += 1
                    else:
                        defaultHealthRulesModified += 1
                analysisDataEvaluatedMetrics["numberOfDefaultHealthRulesModified"] = defaultHealthRulesModified

                # numberOfActionsBoundToEnabledPolicies
                actionsInEnabledPolicies = set()
                for policy in application["policies"]:
                    if policy["enabled"]:
                        for action in policy["actions"]:
                            actionsInEnabledPolicies.add(action["actionName"])
                analysisDataEvaluatedMetrics["numberOfActionsBoundToEnabledPolicies"] = len(actionsInEnabledPolicies)

                # numberOfCustomHealthRules
                analysisDataEvaluatedMetrics["numberOfCustomHealthRules"] = len(
                    set(application["healthRules"].keys()).symmetric_difference(defaultHealthRules.keys())
                )

                analysisDataRawMetrics["totalWarningPolicyViolationsLast24Hours"] = policyEventCounts["warning"]
                analysisDataRawMetrics["totalCriticalPolicyViolationsLast24Hours"] = policyEventCounts["critical"]
                analysisDataRawMetrics["numberOfHealthRules"] = len(application["healthRules"])
                analysisDataRawMetrics["numberOfPolicies"] = len(application["policies"])

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)
