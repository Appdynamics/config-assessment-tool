import json
import logging
from collections import OrderedDict

from backend.api.appd.AppDService import AppDService
from deepdiff import DeepDiff
from backend.extractionSteps.JobStepBase import JobStepBase
from backend.util.asyncio_utils import AsyncioUtils


logger = logging.getLogger(__name__.split('.')[-1])


class HealthRulesAndAlertingAPM(JobStepBase):
    def __init__(self):
        super().__init__("apm")

    async def extract(self, controllerData):
        """
        Extract health rule and alerting configuration details.
        1. Makes one API call per application to get Health Rules.
        2. Makes one API call per application to get Event Counts (health rule violations).
        3. Makes one API call per application to get Policies.
        """
        jobStepName = type(self).__name__

        for host, hostInfo in controllerData.items():
            logger.info(f'{hostInfo["controller"].host} - Extracting {jobStepName}')
            controller: AppDService = hostInfo["controller"]

            # Gather necessary metrics.
            getHealthRulesFutures = []
            getEventCountsFutures = []
            getPoliciesFutures = []
            for application in hostInfo[self.componentType].values():
                getEventCountsFutures.append(
                    controller.getEventCounts(
                        applicationID=application["id"],
                        entityType="APPLICATION",
                        entityID=application["id"],
                    )
                )
                getHealthRulesFutures.append(controller.getHealthRules(application["id"]))
                getPoliciesFutures.append(controller.getPolicies(application["id"]))

            eventCounts = await AsyncioUtils.gatherWithConcurrency(*getEventCountsFutures)
            healthRules = await AsyncioUtils.gatherWithConcurrency(*getHealthRulesFutures)
            policies = await AsyncioUtils.gatherWithConcurrency(*getPoliciesFutures)

            for idx, applicationName in enumerate(hostInfo[self.componentType]):
                application = hostInfo[self.componentType][applicationName]

                if eventCounts[idx].error is not None or not isinstance(eventCounts[idx].data, dict):
                    logger.warning(f"Failed to retrieve event counts for '{applicationName}': {eventCounts[idx].error}; defaulting to zero violations.")
                    application["eventCounts"] = {"policyViolationEventCounts": {"totalPolicyViolations": {"warning": 0, "critical": 0}}}
                else:
                    application["eventCounts"] = eventCounts[idx].data
                application["policies"] = policies[idx].data

                trimmedHrs = [healthRule for healthRule in healthRules[idx].data if healthRule.error is None]
                application["healthRules"] = {
                    healthRuleList.data["name"]: healthRuleList.data for healthRuleList in trimmedHrs if healthRuleList.error is None
                }

    def analyze(self, controllerData, thresholds):
        """
        Analysis of error configuration details.
        1. Determines number of Health Rule violations.
        2. Determines number of Default Health Rules modified.
        3. Determines number of Actions currently bound to enabled policies.
        4. Determines number of Custom Health Rules.
        """

        jobStepName = type(self).__name__

        # Get thresholds related to job
        jobStepThresholds = thresholds[self.componentType][jobStepName]

        defaultHealthRules = json.loads(open("backend/resources/controllerDefaults/defaultHealthRulesAPM.json").read())
        for host, hostInfo in controllerData.items():
            logger.info(f'{hostInfo["controller"].host} - Analyzing {jobStepName}')

            for application in hostInfo[self.componentType].values():
                # Root node of current application for current JobStep.
                analysisDataRoot = application[jobStepName] = OrderedDict()
                # This data goes into the 'JobStep - Metrics' xlsx sheet.
                analysisDataEvaluatedMetrics = analysisDataRoot["evaluated"] = OrderedDict()
                # This data goes into the 'JobStep - Raw' xlsx sheet.
                analysisDataRawMetrics = analysisDataRoot["raw"] = OrderedDict()

                # numberOfHealthRuleViolations
                policyEventCounts = application["eventCounts"]["policyViolationEventCounts"]["totalPolicyViolations"]
                analysisDataEvaluatedMetrics["numberOfHealthRuleViolations"] = policyEventCounts["warning"] + policyEventCounts["critical"]

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
                            logger.debug(f'[{application["name"]}] Default health rule MODIFIED: "{hrName}"')
                            # values_changed: field exists in both but value differs
                            for path, change in healthRuleDiff.get("values_changed", {}).items():
                                logger.debug(
                                    f'  [CHANGED]  {path}\n'
                                    f'             expected (default): {change["old_value"]}\n'
                                    f'             actual   (API):     {change["new_value"]}'
                                )
                            # type_changes: same field, different Python type
                            for path, change in healthRuleDiff.get("type_changes", {}).items():
                                logger.debug(
                                    f'  [TYPE CHG] {path}\n'
                                    f'             expected (default): {change["old_value"]} ({change["old_type"].__name__})\n'
                                    f'             actual   (API):     {change["new_value"]} ({change["new_type"].__name__})'
                                )
                            # dictionary_item_added: fields present in API but not in default
                            for path in healthRuleDiff.get("dictionary_item_added", set()):
                                logger.debug(
                                    f'  [ADDED]    {path} — present in API response but NOT in default'
                                )
                            # dictionary_item_removed: fields in default but missing from API
                            for path in healthRuleDiff.get("dictionary_item_removed", set()):
                                logger.debug(
                                    f'  [REMOVED]  {path} — expected in default but MISSING from API response'
                                )
                            # iterable_item_added / removed — drill into dict items to show field-level diff
                            for path, val in healthRuleDiff.get("iterable_item_added", {}).items():
                                if isinstance(val, dict):
                                    label = val.get("name") or val.get("shortName") or str(val)[:60]
                                    logger.debug(f'  [LIST+]    {path} — item in API but NOT in default: "{label}"')
                                else:
                                    logger.debug(f'  [LIST+]    {path} — item in API but NOT in default: {val}')
                            for path, val in healthRuleDiff.get("iterable_item_removed", {}).items():
                                if isinstance(val, dict):
                                    label = val.get("name") or val.get("shortName") or str(val)[:60]
                                    logger.debug(f'  [LIST-]    {path} — item in default but NOT in API: "{label}"')
                                else:
                                    logger.debug(f'  [LIST-]    {path} — item in default but NOT in API: {val}')
                            # iterable_item_changed — both sides are dicts; drill in to show only differing keys
                            for path, change in healthRuleDiff.get("iterable_item_changed", {}).items():
                                old_val = change.get("old_value", {})
                                new_val = change.get("new_value", {})
                                if isinstance(old_val, dict) and isinstance(new_val, dict):
                                    inner_diff = DeepDiff(old_val, new_val, ignore_order=True)
                                    label = old_val.get("name") or old_val.get("shortName") or path
                                    logger.debug(f'  [LIST CHG] {path} (condition: "{label}")')
                                    for ipath, ichange in inner_diff.get("values_changed", {}).items():
                                        logger.debug(
                                            f'    [CHANGED]  {ipath}\n'
                                            f'               expected (default): {ichange["old_value"]}\n'
                                            f'               actual   (API):     {ichange["new_value"]}'
                                        )
                                    for ipath in inner_diff.get("dictionary_item_added", set()):
                                        logger.debug(f'    [ADDED]    {ipath} — extra field in API (not in default)')
                                    for ipath in inner_diff.get("dictionary_item_removed", set()):
                                        logger.debug(f'    [REMOVED]  {ipath} — missing from API response')
                                else:
                                    logger.debug(
                                        f'  [LIST CHG] {path}\n'
                                        f'             expected (default): {old_val}\n'
                                        f'             actual   (API):     {new_val}'
                                    )
                        else:
                            logger.debug(
                                f'[{application["name"]}] Default health rule unchanged: "{hrName}"'
                            )
                    else:
                        defaultHealthRulesModified += 1
                        logger.debug(
                            f'[{application["name"]}] Default health rule MISSING (not found in application): "{hrName}"'
                        )
                logger.debug(
                    f'[{application["name"]}] Total default health rules modified/missing: {defaultHealthRulesModified} / {len(defaultHealthRules)}'
                )
                analysisDataEvaluatedMetrics["numberOfDefaultHealthRulesModified"] = defaultHealthRulesModified

                # numberOfActionsBoundToEnabledPolicies
                actionsInEnabledPolicies = set()
                for policy in application["policies"]:
                    if policy["enabled"]:
                        if "actions" in policy:
                            for action in policy["actions"]:
                                actionsInEnabledPolicies.add(action["actionName"])
                        else:
                            logger.warning(f"Policy {policy['name']} is enabled but has no actions bound to it.")
                analysisDataEvaluatedMetrics["numberOfActionsBoundToEnabledPolicies"] = len(actionsInEnabledPolicies)

                # numberOfCustomHealthRules
                analysisDataEvaluatedMetrics["numberOfCustomHealthRules"] = len(
                    set(application["healthRules"].keys()).symmetric_difference(defaultHealthRules.keys())
                )

                analysisDataRawMetrics["totalWarningPolicyViolations"] = policyEventCounts["warning"]
                analysisDataRawMetrics["totalCriticalPolicyViolations"] = policyEventCounts["critical"]
                analysisDataRawMetrics["numberOfHealthRules"] = len(application["healthRules"])
                analysisDataRawMetrics["numberOfPolicies"] = len(application["policies"])

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)
