import json
import logging
import re
import time
from datetime import date, datetime, timedelta
from json import JSONDecodeError
from math import ceil
from typing import List

from api.Result import Result
from api.appd.AppDController import AppdController
from api.appd.AuthMethod import AuthMethod
from util.asyncio_utils import AsyncioUtils
from util.stdlib_utils import get_recursively


class AppDService:
    controller: AppdController
    authMethod: AuthMethod

    def __init__(self,
                 applicationFilter: dict = None,
                 timeRangeMins: int = 1440,
                 authMethod: AuthMethod = None):

        self.applicationFilter = applicationFilter
        self.timeRangeMins = timeRangeMins
        self.endTime = int(round(time.time() * 1000))
        self.startTime = self.endTime - (1 * 60 * self.timeRangeMins * 1000)
        self.totalCallsProcessed = 0

        self.authMethod = authMethod
        self.host = authMethod.host
        self.controller = authMethod.controller
        self.username = authMethod.username

    def getAuthMethod(self) -> AuthMethod:
        return self.authMethod

    def __json__(self):
        return {
            "host": self.host,
            "username": self.username,
        }

    async def loginToController(self) -> Result:
        logging.debug(f"{self.host} - Attempt controller connection.")
        try:
            response = await self.controller.login()
        except Exception as e:
            logging.error(f"{self.host} - Controller login failed with {e}")
            return Result(
                None,
                Result.Error(f"{self.host} - {e}."),
            )
        if response.status_code != 200:
            err_msg = f"{self.host} - Controller login failed with " \
                         f"{response.status_code}. Check username and password."
            logging.error(err_msg)
            return Result(
                response,
                Result.Error(err_msg),
            )
        try:
            jsessionid = \
                re.search("JSESSIONID=(\\w|\\d)*", str(response.headers)).group(
                    0).split("JSESSIONID=")[1]
            self.controller.jsessionid = jsessionid
        except AttributeError:
            logging.debug(
                f"{self.host} - Unable to find JSESSIONID in login response. Please verify credentials.")
        try:
            xcsrftoken = \
                re.search("X-CSRF-TOKEN=(\\w|\\d)*",
                          str(response.headers)).group(
                    0).split("X-CSRF-TOKEN=")[1]
            self.controller.xcsrftoken = xcsrftoken
        except AttributeError:
            logging.debug(
                f"{self.host} - Unable to find X-CSRF-TOKEN in login response. Please verify credentials.")

        if self.controller.jsessionid is None or self.controller.xcsrftoken is None:
            return Result(
                response,
                Result.Error(
                    f"{self.host} - Valid authentication headers not cached from previous login call. Please verify credentials."),
            )

        self.controller.session.headers[
            "X-CSRF-TOKEN"] = self.controller.xcsrftoken
        self.controller.session.headers[
            "Set-Cookie"] = f"JSESSIONID={self.controller.jsessionid};X-CSRF-TOKEN={self.controller.xcsrftoken};"
        self.controller.session.headers[
            "Content-Type"] = "application/json;charset=UTF-8"

        logging.debug(f"{self.host} - Controller initialization successful.")
        return Result(self.controller, None)

    async def getBTs(self, applicationID: int) -> Result:
        debugString = f"Gathering bts"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getBTs(applicationID)
        return await self.getResultFromResponse(response, debugString)

    async def getApmApplications(self) -> Result:
        debugString = f"Gathering applications"
        logging.debug(f"{self.host} - {debugString}")

        if self.applicationFilter is not None:
            if self.applicationFilter.get("apm") is None:
                logging.warning(
                    f"Filtered out all APM applications from analysis by match rule {self.applicationFilter['apm']}")
                return Result([], None)

        response = await self.controller.getApmApplications()
        result = await self.getResultFromResponse(response, debugString)
        # apparently it's possible to have a null application name, the controller converts the null into "null"
        if result.error is None:
            for application in result.data:
                if application["name"] is None:
                    application["name"] = "null"

        if self.applicationFilter is not None:
            pattern = re.compile(self.applicationFilter["apm"])
            for application in result.data:
                if not pattern.search(application["name"]):
                    logging.warning(
                        f"Filtered out APM application {application['name']} from analysis by match rule {self.applicationFilter['apm']}")
            result.data = [application for application in result.data if
                           pattern.search(application["name"])]

        return result

    async def getNode(self, applicationID: int, nodeID: int) -> Result:
        debugString = f"Getting single node for Application:{applicationID} node:{nodeID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getNode(applicationID, nodeID)
        return await self.getResultFromResponse(response, debugString)

    async def getNodes(self, applicationID: int) -> Result:
        debugString = f"Gathering nodes for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getNodes(applicationID)
        return await self.getResultFromResponse(response, debugString)

    async def getTiers(self, applicationID: int) -> Result:
        debugString = f"Gathering tiers for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getTiers(applicationID)
        return await self.getResultFromResponse(response, debugString)

    async def getBtMatchRules(self, applicationID: int) -> Result:
        debugString = f"Gathering Application Business Transaction Custom Match Rules for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getBtMatchRules(applicationID)
        return await self.getResultFromResponse(response, debugString)

    async def getBackends(self, applicationID: int) -> Result:
        debugString = f"Gathering Backends for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getBackends(applicationID)
        return await self.getResultFromResponse(response, debugString)

    async def getConfigurations(self) -> Result:
        debugString = f"Gathering Controller Configurations"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getConfigurations()
        return await self.getResultFromResponse(response, debugString)

    # TODO: need to look at individual tiers as well, and individual agentTypes
    async def getAllCustomExitPoints(self, applicationID: int) -> Result:
        debugString = f"Gathering Custom Exit Points for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        body = '{"agentType": "APP_AGENT", "attachedEntity": {"entityId": {applicationID}, "entityType": "APPLICATION"}}'.replace(
            "{applicationID}", str(applicationID)
        )
        response = await self.controller.getAllCustomExitPoints(body)
        return await self.getResultFromResponse(response, debugString)

    # TODO: need to look at individual tiers as well, and individual agent types
    async def getBackendDiscoveryConfigs(self, applicationID: int) -> Result:
        debugString = f"Gathering Backend Discovery Configs for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        body = '{"agentType": "APP_AGENT", "attachedEntity": {"entityId": {applicationID}, "entityType": "APPLICATION"}}'.replace(
            "{applicationID}", str(applicationID)
        )
        response = await self.controller.getBackendDiscoveryConfigs(body)
        return await self.getResultFromResponse(response, debugString)

    async def getDevModeConfig(self, applicationID: int) -> Result:
        debugString = f"Gathering Developer Mode Configuration for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getDevModeConfig(applicationID)
        return await self.getResultFromResponse(response, debugString)

    async def getInstrumentationLevel(self, applicationID: int) -> Result:
        debugString = f"Gathering Instrumentation Level for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getInstrumentationLevel(applicationID)
        return await self.getResultFromResponse(response, debugString,
                                                isResponseJSON=False)

    async def getAllNodePropertiesForCustomizedComponents(self,
                                                          applicationID: int) -> Result:
        debugString = f"Gathering All Application Components With Nodes for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getAllApplicationComponentsWithNodes(
            applicationID)
        applicationComponentsWithNodes = await self.getResultFromResponse(
            response, debugString)

        getAgentConfigurationFutures = []
        for applicationConfiguration in applicationComponentsWithNodes.data:
            if applicationConfiguration["children"] is not None:
                # Application is always customized, need to fetch both java and .NET application config
                getAgentConfigurationFutures.append(
                    self.getAgentConfiguration(
                        applicationID,
                        "APP_AGENT",
                        applicationConfiguration["entityType"],
                        applicationConfiguration["agentConfigId"],
                    )
                )
                getAgentConfigurationFutures.append(
                    self.getAgentConfiguration(
                        applicationID,
                        "DOT_NET_APP_AGENT",
                        applicationConfiguration["entityType"],
                        applicationConfiguration["agentConfigId"],
                    )
                )
                for tierConfiguration in applicationConfiguration["children"]:
                    if tierConfiguration["children"] is not None:
                        if tierConfiguration["customized"]:
                            getAgentConfigurationFutures.append(
                                self.getAgentConfiguration(
                                    applicationID,
                                    tierConfiguration["agentType"],
                                    tierConfiguration["entityType"],
                                    tierConfiguration["agentConfigId"],
                                )
                            )
                        for nodeConfiguration in tierConfiguration["children"]:
                            if nodeConfiguration["customized"]:
                                getAgentConfigurationFutures.append(
                                    self.getAgentConfiguration(
                                        applicationID,
                                        nodeConfiguration["agentType"],
                                        nodeConfiguration["entityType"],
                                        nodeConfiguration["agentConfigId"],
                                    )
                                )
        # TODO: this needs to be batched
        agentConfigurations = await AsyncioUtils.gatherWithConcurrency(*getAgentConfigurationFutures)
        return Result(agentConfigurations, None)

    async def getAgentConfiguration(self, applicationID: int, agentType: str,
                                    entityType: str, entityId: int) -> Result:
        debugString = f"Gathering Agent Configuration for Application:{applicationID} entity:{entityId}"
        logging.debug(f"{self.host} - {debugString}")
        body = (
            '{"checkAncestors":false,"key":{"agentType":"{agentType}","attachedEntity":{"id":null,"version":null,"entityId":{entityId},"entityType":"{entityType}"}}}'.replace(
                "{agentType}", str(agentType)
            )
            .replace("{entityId}", str(entityId))
            .replace("{entityType}", entityType)
        )
        response = await self.controller.getAgentConfiguration(body)
        return await self.getResultFromResponse(response, debugString)

    async def getApplicationConfiguration(self, applicationID: int) -> Result:
        debugString = f"Gathering Application Call Graph Settings for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getApplicationConfiguration(
            applicationID)
        return await self.getResultFromResponse(response, debugString,
                                                isResponseList=False)

    async def getServiceEndpointMatchRules(self, applicationID: int) -> Result:
        debugString = f"Gathering Service Endpoint Custom Match Rules for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getApplicationComponents(applicationID)
        response = await self.getResultFromResponse(response, debugString)

        customMatchRulesFutures = []
        defaultMatchRulesFutures = []
        body = '{"agentType":"APP_AGENT","attachedEntity":{"entityId":{entityId},"entityType":"APPLICATION"}}'.replace(
            "{entityId}", f"{applicationID}"
        )
        defaultMatchRulesFutures.append(
            self.controller.getServiceEndpointDefaultMatchRules(body))
        for entity in response.data:
            body = '{"attachedEntity":{"entityType":"APPLICATION_COMPONENT","entityId":{entityId}},"agentType":"{agentType}"}'.replace(
                "{entityId}", str(entity["id"])
            ).replace("{agentType}", str(entity["componentType"]["agentType"]))
            customMatchRulesFutures.append(
                self.controller.getServiceEndpointCustomMatchRules(body))

            body = '{"agentType":"APP_AGENT","attachedEntity":{"entityId":{entityId},"entityType":"APPLICATION_COMPONENT"}}'.replace(
                "{entityId}", str(entity["id"])
            )
            defaultMatchRulesFutures.append(
                self.controller.getServiceEndpointDefaultMatchRules(body))

        response = await AsyncioUtils.gatherWithConcurrency(
            *customMatchRulesFutures)
        customMatchRules = [
            await self.getResultFromResponse(response, debugString) for response
            in response]
        response = await AsyncioUtils.gatherWithConcurrency(
            *defaultMatchRulesFutures)
        defaultMatchRules = [
            await self.getResultFromResponse(response, debugString) for response
            in response]

        return Result((customMatchRules, defaultMatchRules), None)

    async def getAppLevelBTConfig(self, applicationID: int) -> Result:
        debugString = f"Gathering Application Business Transaction Configuration Settings for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getAppLevelBTConfig(applicationID)
        return await self.getResultFromResponse(response, debugString)

    async def getCustomMetrics(self, applicationID: int,
                               tierName: str) -> Result:
        debugString = f"Gathering Custom Metrics for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "request": None,
            "applicationId": applicationID,
            "livenessStatus": "ALL",
            "pathData": ["Application Infrastructure Performance", tierName,
                         "Custom Metrics"],
            "timeRangeSpecifier": {
                "type": "BEFORE_NOW",
                "durationInMinutes": self.timeRangeMins,
                "endTime": None,
                "startTime": None,
                "timeRange": None,
                "timeRangeAdjusted": False,
            },
        }
        response = await self.controller.getMetricTree(json.dumps(body))
        return await self.getResultFromResponse(response, debugString)

    async def getMetricData(
            self,
            applicationID: int,
            metric_path: str,
            rollup: bool,
            time_range_type: str,
            duration_in_mins: int = "",
            start_time: int = "",
            end_time: int = 1440,
    ) -> Result:
        debugString = f'Gathering Metrics for:"{metric_path}" on application:{applicationID}'
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getMetricData(
            applicationID,
            metric_path,
            rollup,
            time_range_type,
            duration_in_mins,
            start_time,
            end_time,
        )
        return await self.getResultFromResponse(response, debugString)

    async def getApplicationEvents(
            self,
            applicationID: int,
            event_types: List[str],
            severities: List[str],
            time_range_type: str,
            duration_in_mins: str = "",
            start_time: str = "",
            end_time: str = "",
    ) -> Result:
        debugString = f'Gathering Application Events for:"{event_types}" with severities {severities} on application:{applicationID}'
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getApplicationEvents(
            applicationID,
            ",".join(event_types),
            ",".join(severities),
            time_range_type,
            duration_in_mins,
            start_time,
            end_time,
        )
        return await self.getResultFromResponse(response, debugString)

    async def getEventCounts(self, applicationID: int, entityType: str,
                             entityID: int) -> Result:
        debugString = f'Gathering Event Counts for:"{entityType}" {entityID} on application:{applicationID}'
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getEventCounts(
            applicationID, entityType, entityID,
            f"Custom_Time_Range.BETWEEN_TIMES.{self.endTime}.{self.startTime}.{self.timeRangeMins}"
        )
        return await self.getResultFromResponse(response, debugString)

    async def getHealthRules(self, applicationID: int) -> Result:
        debugString = f"Gathering Health Rules for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getHealthRules(applicationID)
        healthRules = await self.getResultFromResponse(response, debugString)

        healthRuleDetails = []
        for healthRule in healthRules.data:
            healthRuleDetails.append(
                self.controller.getHealthRule(applicationID, healthRule["id"]))

        responses = await AsyncioUtils.gatherWithConcurrency(*healthRuleDetails)

        healthRulesData = []
        for response, healthRule in zip(responses, healthRules.data):
            debugString = f"Gathering Health Rule Data for Application:{applicationID} HealthRule:'{healthRule['name']}'"
            healthRulesData.append(
                await self.getResultFromResponse(response, debugString))

        return Result(healthRulesData, None)

    async def getPolicies(self, applicationID: int) -> Result:
        debugString = f"Gathering Policies for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getPolicies(applicationID)
        return await self.getResultFromResponse(response, debugString)

    async def getSnapshotsWithDataCollector(
            self,
            applicationID: int,
            data_collector_name: str,
            data_collector_type: str,
            maximum_results: int = 1,
            data_collector_value: str = "",
    ) -> Result:
        debugString = f"Gathering Snapshots for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "firstInChain": False,
            "maxRows": maximum_results,
            "applicationIds": [applicationID],
            "businessTransactionIds": [],
            "applicationComponentIds": [],
            "applicationComponentNodeIds": [],
            "errorIDs": [],
            "errorOccured": None,
            "userExperience": [],
            "executionTimeInMilis": None,
            "endToEndLatency": None,
            "url": None,
            "sessionId": None,
            "userPrincipalId": None,
            "dataCollectorFilter": {"collectorType": data_collector_type,
                                    "query": {"name": data_collector_name,
                                              "value": ""}},
            "archived": None,
            "guids": [],
            "diagnosticSnapshot": None,
            "badRequest": None,
            "deepDivePolicy": [],
            "rangeSpecifier": {"type": "BEFORE_NOW",
                               "durationInMinutes": self.timeRangeMins},
        }
        response = await self.controller.getSnapshotsWithDataCollector(
            json.dumps(body))

        return await self.getResultFromResponse(response, debugString)

    async def getDataCollectorUsage(self, applicationID: int) -> Result:
        debugString = f"Gathering Data Collectors for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getDataCollectors(applicationID)

        dataCollectors = await self.getResultFromResponse(response, debugString)
        snapshotEnabledDataCollectors = [dataCollector for dataCollector in
                                         dataCollectors.data if
                                         dataCollector["enabledForApm"]]

        httpDataCollectors = [dataCollector for dataCollector in
                              snapshotEnabledDataCollectors if
                              dataCollector["type"] == "http"]
        dataCollectorFields = []
        for dataCollector in httpDataCollectors:
            for field in dataCollector["requestParameters"]:
                dataCollectorFields.append(
                    (
                        "HTTP Parameter",
                        field["displayName"],
                        dataCollector["enabledForAnalytics"],
                    )
                )
            for field in dataCollector["cookieNames"]:
                dataCollectorFields.append(
                    ("Cookie", field, dataCollector["enabledForAnalytics"]))
            for field in dataCollector["sessionKeys"]:
                dataCollectorFields.append(("Session Key", field, dataCollector[
                    "enabledForAnalytics"]))
            for field in dataCollector["headers"]:
                dataCollectorFields.append(("HTTP Header", field, dataCollector[
                    "enabledForAnalytics"]))

        pojoDataCollectors = [dataCollector for dataCollector in
                              snapshotEnabledDataCollectors if
                              dataCollector["type"] == "pojo"]
        for dataCollector in pojoDataCollectors:
            for methodDataGathererConfig in dataCollector[
                "methodDataGathererConfigs"]:
                dataCollectorFields.append(
                    (
                        "Business Data",
                        methodDataGathererConfig["name"],
                        dataCollector["enabledForAnalytics"],
                    )
                )

        snapshotsContainingDataCollectorFields = []
        distinctDataCollectors = set()
        for dataCollectorField in dataCollectorFields:
            if (applicationID, dataCollectorField[1],
                dataCollectorField[0]) not in distinctDataCollectors:
                snapshotsContainingDataCollectorFields.append(
                    self.getSnapshotsWithDataCollector(
                        applicationID=applicationID,
                        data_collector_name=dataCollectorField[1],
                        data_collector_type=dataCollectorField[0],
                    )
                )
            distinctDataCollectors.add(
                (applicationID, dataCollectorField[1], dataCollectorField[0]))
        snapshotResults = await AsyncioUtils.gatherWithConcurrency(
            *snapshotsContainingDataCollectorFields)

        dataCollectorFieldsWithSnapshots = []
        for collector, snapshotResult in zip(dataCollectorFields,
                                             snapshotResults):
            if snapshotResult.error is None and len(
                    snapshotResult.data["requestSegmentDataListItems"]) == 1:
                dataCollectorFieldsWithSnapshots.append(collector)
            # This API does not work for either session keys or headers, as far as I know there is no way to get this info without inspecting ALL snapshots (won't do).
            # The API comes from the Transaction Snapshot filtering UI. No UI option for session keys or headers exists there.
            # For now, let's just assume that any session key or header configured data collector is working... If anyone has a better idea I'm all ears.
            elif collector[0] == "Session Key" or collector[0] == "HTTP Header":
                dataCollectorFieldsWithSnapshots.append(collector)

        result = {
            "allDataCollectors": dataCollectorFields,
            "dataCollectorsPresentInSnapshots": dataCollectorFieldsWithSnapshots,
            "dataCollectorsPresentInAnalytics": [dataCollector for dataCollector
                                                 in
                                                 dataCollectorFieldsWithSnapshots
                                                 if dataCollector[2]],
        }
        return Result(result, None)

    async def getAnalyticsEnabledStatusForAllApplications(self) -> Result:
        debugString = f"Gathering Analytics Enabled Status for all Applications"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getAnalyticsEnabledStatusForAllApplications()
        return await self.getResultFromResponse(response, debugString)

    async def getDashboards(self) -> Result:

        debugString = f"Gathering Dashboards"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getAllDashboardsMetadata()
        allDashboardsMetadata = await self.getResultFromResponse(response,
                                                                 debugString)

        dashboards = []
        batch_size = AsyncioUtils.concurrentConnections
        for i in range(0, len(allDashboardsMetadata.data), batch_size):
            dashboardsFutures = []

            logging.debug(
                f"Batch iteration {int(i / batch_size)} of {ceil(len(allDashboardsMetadata.data) / batch_size)}")
            chunk = allDashboardsMetadata.data[i: i + batch_size]

            for dashboard in chunk:
                dashboardsFutures.append(
                    self.controller.getDashboard(dashboard["id"]))

            response = await AsyncioUtils.gatherWithConcurrency(
                *dashboardsFutures)
            for dashboard in [
                await self.getResultFromResponse(response, debugString) for
                response in response]:
                dashboards.append(dashboard.data)
                # logging.info(f'{self.host} - DASHBOARD: '
                #              f'{dashboard.data["name"]}'
                #              f' {dashboard.data["name"]}')

        # The above implementation shouldn't be necessary since gatherWithConcurrency uses a semaphore to limit number of concurrent calls.
        # But on controllers with a large number of dashboards the coroutines will get stuck unless explicitly batched.
        # The below implementation should work, but doesn't and I'm tired of looking at it. Maybe someone smarter than me can fix it.
        # Tens of hours wasted here, beware ye who enter.
        # dashboardFutures = []
        # for dashboard in allDashboardsMetadata.data:
        #     dashboardFutures.append(self.controller.getDashboard(dashboard["id"]))
        # response = await gatherWithConcurrency(*dashboardFutures)
        # dashboards = [await self.getResultFromResponse(response, debugString) for response in response]
        # dashboards = [dashboard.data for dashboard in dashboards if dashboard.error is None]

        returnedDashboards = []
        for dashboardSchema, dashboardOverview in zip(dashboards,
                                                      allDashboardsMetadata.data):
            if "schemaVersion" in dashboardSchema:
                dashboardSchema["createdBy"] = dashboardOverview["createdBy"]
                dashboardSchema["createdOn"] = dashboardOverview["createdOn"]
                dashboardSchema["modifiedOn"] = dashboardOverview["modifiedOn"]
                returnedDashboards.append(dashboardSchema)

        return Result(returnedDashboards, None)

    async def getUserPermissions(self, username: str) -> Result:
        debugString = f"Gathering Permission set for user: {username}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.getAuthMethod().validatePermissions()

        # response = await self.controller.getUsers()
        # users = await self.getResultFromResponse(response, debugString)

        if response.error is not None:
            logging.error(
                f"{self.host} - Call to Get User Permissions failed. Is user '{self.username}' an Account Owner?")
            return Result(
                response,
                Result.Error(
                    f"{self.host} - Call to Get User Permissions failed. Is user '{self.username}' an Account Owner?"),
            )

        return await self.getResultFromResponse(response, debugString)

    async def getAccountUsageSummary(self) -> Result:
        debugString = f"Gathering Account Usage Summary"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "type": "BEFORE_NOW",
            "durationInMinutes": self.timeRangeMins,
            "endTime": None,
            "startTime": None,
            "timeRange": None,
            "timeRangeAdjusted": False,
        }
        response = await self.controller.getAccountUsageSummary(
            json.dumps(body))
        return await self.getResultFromResponse(response, debugString)

    async def getEumLicenseUsage(self) -> Result:
        debugString = f"Gathering Account Usage Summary"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "type": "BEFORE_NOW",
            "durationInMinutes": self.timeRangeMins,
            "endTime": None,
            "startTime": None,
            "timeRange": None,
            "timeRangeAdjusted": False,
        }
        response = await self.controller.getEumLicenseUsage(json.dumps(body))
        return await self.getResultFromResponse(response, debugString)

    async def getAppAgentMetadata(self, applicationId: int,
                                  agentIDs: list[str]) -> Result:
        debugString = f"Gathering App Agent Metadata"
        logging.debug(f"{self.host} - {debugString}")

        if len(agentIDs) == 0:
            return Result([], None)

        futures = [
            self.controller.getAppServerAgentsMetadata(applicationId, agentId)
            for agentId in agentIDs]
        response = await AsyncioUtils.gatherWithConcurrency(*futures)
        results = [
            (await self.getResultFromResponse(response, debugString)).data for
            response in response]
        return Result(results, None)

    async def getAppServerAgents(self) -> Result:
        debugString = f"Gathering App Server Agents Agents"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "requestFilter": {
                "queryParams": {"applicationAssociationType": "ALL"},
                "filters": []},
            "resultColumns": [],
            "offset": 0,
            "limit": -1,
            "searchFilters": [],
            "columnSorts": [{"column": "HOST_NAME", "direction": "ASC"}],
            "timeRangeStart": self.startTime,
            "timeRangeEnd": self.endTime,
        }
        response = await self.controller.getAppServerAgents(json.dumps(body))
        result = await self.getResultFromResponse(response, debugString)
        if result.error is not None:
            return result

        agentIds = [agent["applicationComponentNodeId"] for agent in
                    result.data["data"]]

        debugString = f"Gathering App Server Agents Agents List"
        agentFutures = []
        batch_size = 50
        for i in range(0, len(agentIds), batch_size):
            chunk = agentIds[i: i + batch_size]
            body = {
                "requestFilter": chunk,
                "resultColumns": [
                    "HOST_NAME",
                    "AGENT_VERSION",
                    "NODE_NAME",
                    "COMPONENT_NAME",
                    "APPLICATION_NAME",
                    "DISABLED",
                    "ALL_MONITORING_DISABLED",
                ],
                "offset": 0,
                "limit": -1,
                "searchFilters": [],
                "columnSorts": [{"column": "HOST_NAME", "direction": "ASC"}],
                "timeRangeStart": self.startTime,
                "timeRangeEnd": self.endTime,
            }
            agentFutures.append(
                self.controller.getAppServerAgentsIds(json.dumps(body)))

        response = await AsyncioUtils.gatherWithConcurrency(*agentFutures)
        results = [
            (await self.getResultFromResponse(response, debugString)).data[
                "data"] for response in response]
        out = []
        for result in results:
            out.extend(result)
        return Result(out, None)

    async def getMachineAgents(self) -> Result:
        debugString = f"Gathering App Server Agents Agents"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "requestFilter": {
                "queryParams": {"applicationAssociationType": "ALL"},
                "filters": []},
            "resultColumns": [],
            "offset": 0,
            "limit": -1,
            "searchFilters": [],
            "columnSorts": [{"column": "HOST_NAME", "direction": "ASC"}],
            "timeRangeStart": self.startTime,
            "timeRangeEnd": self.endTime,
        }
        response = await self.controller.getMachineAgents(json.dumps(body))
        result = await self.getResultFromResponse(response, debugString)
        if result.error is not None:
            return result

        agentIds = [agent["machineId"] for agent in result.data["data"]]

        debugString = f"Gathering Machine Agents Agents List"
        agentFutures = []
        batch_size = 50
        for i in range(0, len(agentIds), batch_size):
            chunk = agentIds[i: i + batch_size]
            body = {
                "requestFilter": chunk,
                "resultColumns": ["AGENT_VERSION", "APPLICATION_NAMES",
                                  "ENABLED"],
                "offset": 0,
                "limit": -1,
                "searchFilters": [],
                "columnSorts": [{"column": "HOST_NAME", "direction": "ASC"}],
                "timeRangeStart": self.startTime,
                "timeRangeEnd": self.endTime,
            }

            agentFutures.append(
                self.controller.getMachineAgentsIds(json.dumps(body)))

        response = await AsyncioUtils.gatherWithConcurrency(*agentFutures)
        results = [
            (await self.getResultFromResponse(response, debugString)).data[
                "data"] for response in response]
        out = []
        for result in results:
            out.extend(result)
        return Result(out, None)

    async def getDBAgents(self) -> Result:
        debugString = f"Gathering DB Agents"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getDBAgents()
        return await self.getResultFromResponse(response, debugString)

    async def getAnalyticsAgents(self) -> Result:
        debugString = f"Gathering Analytics Agents"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getAnalyticsAgents()
        return await self.getResultFromResponse(response, debugString)

    async def getServers(self) -> Result:
        debugString = f"Gathering Server Keys"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "filter": {
                "appIds": [],
                "nodeIds": [],
                "tierIds": [],
                "types": ["PHYSICAL", "CONTAINER_AWARE"],
                "timeRangeStart": self.startTime,
                "timeRangeEnd": self.endTime,
            },
            "sorter": {"field": "HEALTH", "direction": "ASC"},
        }

        response = await self.controller.getServersKeys(json.dumps(body))
        serverKeys = await self.getResultFromResponse(response, debugString)

        machineIds = []
        if not isinstance(serverKeys.data["machineKeys"], list):
            logging.warning("Expected 'serverKeys.data[\"machineKeys\"]' to be a "
                            "list, but got {}".format(type(serverKeys.data["machineKeys"])))
        else:
            for serverKey in serverKeys.data["machineKeys"]:
                if isinstance(serverKey, dict) and "machineId" in serverKey:
                    try:
                        machineIds.append(serverKey["machineId"])
                    except TypeError:
                        logging.warning("TypeError encountered with "
                                        "machineId: {}".format(serverKey["machineId"]))
                else:
                    if isinstance(serverKey, dict):
                        logging.warning("Dictionary lacks 'machineId' key: {}".format(serverKey))
                    else:
                        logging.warning("Expected a dictionary, but found type: {}".format(type(serverKey)))


        serverFutures = [self.controller.getServer(serverId) for serverId in
                         machineIds]
        serversResponses = await AsyncioUtils.gatherWithConcurrency(
            *serverFutures)

        serverAvailabilityFutures = []
        for machineId in machineIds:
            body = {
                "timeRange": f"Custom_Time_Range.BETWEEN_TIMES.{self.endTime}.{self.startTime}.{self.timeRangeMins}",
                "metricNames": ["Hardware Resources|Machine|Availability"],
                "rollups": [1],
                "ids": [machineId],
                "baselineId": None,
            }
            serverAvailabilityFutures.append(
                self.controller.getServerAvailability(json.dumps(body)))
        serversAvailabilityResponses = await AsyncioUtils.gatherWithConcurrency(
            *serverAvailabilityFutures)

        debugString = f"Gathering Machine Agents Agents List"
        serversResults = [
            (await self.getResultFromResponse(serversResponse, debugString)) for
            serversResponse in serversResponses]
        serversAvailabilityResults = [
            (await self.getResultFromResponse(serversAvailabilityResponse,
                                              debugString))
            for serversAvailabilityResponse in serversAvailabilityResponses
        ]

        machineIdMap = {}
        for serverResult, serverAvailabilityResult in zip(serversResults,
                                                          serversAvailabilityResults):
            machine = serverResult.data
            value = get_recursively(serverAvailabilityResult.data["data"],
                                    "value")
            if value:
                availability = next(iter(value))
                machine["availability"] = availability
            else:
                machine["availability"] = 0

            physicalCores = 0
            virtualCores = 0
            for cpu in machine.get("cpus", []):
                physicalCores += cpu.get("coreCount", 0)
                virtualCores += cpu.get("logicalCount", 0)
            machine["physicalCores"] = physicalCores
            machine["virtualCores"] = virtualCores

            machineIdMap[machine["hostId"]] = machine

        return Result(machineIdMap, None)

    async def getEumApplications(self) -> Result:
        debugString = f"Gathering BRUM Applications"
        logging.debug(f"{self.host} - {debugString}")

        if self.applicationFilter is not None:
            if self.applicationFilter.get("brum") is None:
                logging.warning(
                    f"Filtered out all BRUM applications from analysis by match rule {self.applicationFilter['brum']}")
                return Result([], None)

        response = await self.controller.getEumApplications(
            f"Custom_Time_Range.BETWEEN_TIMES.{self.endTime}.{self.startTime}.{self.timeRangeMins}")
        result = await self.getResultFromResponse(response, debugString)

        if self.applicationFilter is not None:
            pattern = re.compile(self.applicationFilter["brum"])
            for application in result.data:
                if not pattern.search(application["name"]):
                    logging.warning(
                        f"Filtered out BRUM application {application['name']} from analysis by match rule {self.applicationFilter['brum']}"
                    )
            result.data = [application for application in result.data if
                           pattern.search(application["name"])]

        return result

    async def getEumPageListViewData(self, applicationId: int) -> Result:
        debugString = f"Gathering EUM Page List View Data for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "applicationId": applicationId,
            "addId": None,
            "timeRangeString": f"Custom_Time_Range|BETWEEN_TIMES|{self.endTime}|{self.startTime}|{self.timeRangeMins}",
            "fetchSyntheticData": False,
        }
        response = await self.controller.getEumPageListViewData(
            json.dumps(body))
        return await self.getResultFromResponse(response, debugString)

    async def getEumNetworkRequestList(self, applicationId: int) -> Result:
        debugString = f"Gathering EUM Page List View Data for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "requestFilter": {"applicationId": applicationId,
                              "fetchSyntheticData": False},
            "resultColumns": ["PAGE_TYPE", "PAGE_NAME", "TOTAL_REQUESTS",
                              "END_USER_RESPONSE_TIME",
                              "VISUALLY_COMPLETE_TIME"],
            "offset": 0,
            "limit": -1,
            "searchFilters": [],
            "columnSorts": [{"column": "TOTAL_REQUESTS", "direction": "DESC"}],
            "timeRangeStart": self.startTime,
            "timeRangeEnd": self.endTime,
        }
        response = await self.controller.getEumNetworkRequestList(
            json.dumps(body))
        return await self.getResultFromResponse(response, debugString)

    async def getPagesAndFramesConfig(self, applicationId: int) -> Result:
        debugString = f"Gathering Pages and Frames Config for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getPagesAndFramesConfig(applicationId)
        return await self.getResultFromResponse(response, debugString)

    async def getAJAXConfig(self, applicationId: int) -> Result:
        debugString = f"Gathering AJAX Config for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getAJAXConfig(applicationId)
        return await self.getResultFromResponse(response, debugString)

    async def getVirtualPagesConfig(self, applicationId: int) -> Result:
        debugString = f"Gathering Virtual Pages Config for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getVirtualPagesConfig(applicationId)
        return await self.getResultFromResponse(response, debugString)

    async def getBrowserSnapshotsWithServerSnapshots(self,
                                                     applicationId: int) -> Result:
        debugString = f"Gathering Browser Snapshots for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "applicationId": applicationId,
            "timeRangeString": f"Custom_Time_Range.BETWEEN_TIMES.{self.endTime}.{self.startTime}.{self.timeRangeMins}",
            "filters": {
                "_classType": "BrowserSnapshotFilters",
                "serverSnapshotExists": {"type": "BOOLEAN",
                                         "name": "ms_serverSnapshotExists",
                                         "value": True},
                "pages": {
                    "type": "FLY_OUT_SELECT",
                    "name": "ms_pagesAndAjaxRequestsNavLabel",
                    "values": None,
                    "alternateOptionsString": None,
                    "flyoutTitle": "Select Pages",
                },
                "devices": {
                    "type": "FLY_OUT_SELECT",
                    "name": "ms_devices",
                    "values": None,
                    "alternateOptionsString": "Show Only Device Type (mobile vs computer, etc)",
                    "flyoutTitle": "Select Devices",
                },
                "browsers": {
                    "type": "FLY_OUT_SELECT",
                    "name": "ms_browsers",
                    "values": None,
                    "alternateOptionsString": "Show Browser Versions",
                    "flyoutTitle": "Select Browsers",
                },
            },
        }
        response = await self.controller.getBrowserSnapshots(json.dumps(body))
        return await self.getResultFromResponse(response, debugString,
                                                isResponseList=False)

    async def getMRUMApplications(self) -> Result:
        debugString = f"Gathering MRUM Applications"
        logging.debug(f"{self.host} - {debugString}")

        if self.applicationFilter is not None:
            if self.applicationFilter.get("mrum") is None:
                logging.warning(
                    f"Filtered out all MRUM applications from analysis by match rule {self.applicationFilter['mrum']}")
                return Result([], None)

        response = await self.controller.getMRUMApplications(
            f"Custom_Time_Range.BETWEEN_TIMES.{self.endTime}.{self.startTime}.{self.timeRangeMins}")
        result = await self.getResultFromResponse(response, debugString)

        tempData = result.data.copy()
        result.data.clear()
        for mrumApplicationGroup in tempData:
            for mrumApplication in mrumApplicationGroup["children"]:
                mrumApplication["name"] = mrumApplication["internalName"]
                mrumApplication[
                    "taggedName"] = f"{mrumApplicationGroup['appKey']}-{mrumApplication['name']}"
                result.data.append(mrumApplication)

        if self.applicationFilter is not None:
            pattern = re.compile(self.applicationFilter["mrum"])
            for application in result.data:
                if not pattern.search(application["name"]):
                    logging.warning(
                        f"Filtered out MRUM application {application['name']} from analysis by match rule {self.applicationFilter['mrum']}"
                    )
            result.data = [application for application in result.data if
                           pattern.search(application["name"])]

        return result

    async def getMRUMNetworkRequestConfig(self, applicationId: int) -> Result:
        debugString = f"Gathering MRUM Network Request Config for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getMRUMNetworkRequestConfig(
            applicationId)
        return await self.getResultFromResponse(response, debugString)

    async def getNetworkRequestLimit(self, applicationId: int) -> Result:
        debugString = f"Gathering MRUM Network Request Limit for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getNetworkRequestLimit(applicationId)
        return await self.getResultFromResponse(response, debugString)

    async def getMobileSnapshotsWithServerSnapshots(self, applicationId: int,
                                                    mobileApplicationId: int,
                                                    platform: str) -> Result:
        debugString = f"Gathering Mobile Snapshots for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "applicationId": applicationId,
            "timeRangeString": f"Custom_Time_Range|BETWEEN_TIMES|{self.endTime}|{self.startTime}|{self.timeRangeMins}",
            "platform": platform,
            "mobileAppId": mobileApplicationId,
            "serverSnapshotExists": True,
        }
        response = await self.controller.getMobileSnapshots(json.dumps(body))
        return await self.getResultFromResponse(response, debugString)

    async def getSyntheticJobs(self, applicationId: int):
        debugString = f"Gathering Mobile Snapshots for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getSyntheticJobs(applicationId)
        return await self.getResultFromResponse(response, debugString)

    async def getSyntheticBillableTime(self, applicationId: int,
                                       scheduleIds: List[str]) -> Result:
        debugString = f"Gathering Synthetic Billable Time for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "scheduleIds": scheduleIds,
            "appId": applicationId,
            "startTime": self.startTime,
            "currentTime": self.endTime,
        }
        response = await self.controller.getSyntheticBillableTime(
            json.dumps(body))
        return await self.getResultFromResponse(response, debugString)

    async def getSyntheticPrivateAgentUtilization(self, applicationId: int,
                                                  jobsJson: List[
                                                      dict]) -> Result:
        debugString = f"Gathering Synthetic Private Agent Utilization for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getSyntheticPrivateAgentUtilization(
            applicationId, json.dumps(jobsJson))
        return await self.getResultFromResponse(response, debugString)

    async def getSyntheticSessionData(self, applicationId: int,
                                      jobsJson: List[dict]) -> Result:
        debugString = f"Gathering Synthetic Session Data for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        # get the last 24 hours in milliseconds
        lastMonth = self.endTime - (1 * 60 * 60 * 24 * 30 * 1000)
        monthStart = datetime.timestamp(
            datetime.today().replace(day=1, hour=0, minute=0, second=0,
                                     microsecond=0))
        ws = (date.today() - timedelta(date.today().weekday()))
        weekStart = datetime.timestamp(
            datetime.today().replace(year=ws.year, month=ws.month, day=ws.day,
                                     hour=0, minute=0, second=0, microsecond=0))
        body = {
            "appId": applicationId,
            "scheduleIds": jobsJson,
            "fields": ["AVG_DURATION"],
            "timestamps": {
                "startTime": lastMonth,
                "endTime": self.endTime,
                "currentTime": self.endTime,
                "monthStartTime": monthStart,
                "weekStartTime": weekStart,
                "utcOffset": -14400000,
            },
        }
        response = await self.controller.getSyntheticSessionData(
            json.dumps(body))
        return await self.getResultFromResponse(response, debugString)

    async def close(self):
        logging.debug(f"{self.host} - Closing connection")
        await self.authMethod.cleanup()

    async def getResultFromResponse(self, response, debugString,
                                    isResponseJSON=True,
                                    isResponseList=True) -> Result:
        body = (await response.content.read()).decode("ISO-8859-1")
        self.totalCallsProcessed += 1

        if response.status_code >= 400:
            msg = f"{self.host} - {debugString} failed with code:{response.status_code} body:{body}"
            try:
                responseJSON = json.loads(body)
                if "message" in responseJSON:
                    msg = f"{self.host} - {debugString} failed with code:{response.status_code} body:{responseJSON['message']}"
            except JSONDecodeError:
                pass
            logging.debug(msg)
            return Result([] if isResponseList else {},
                          Result.Error(f"{response.status_code}"))
        if isResponseJSON:
            try:
                return Result(json.loads(body), None)
            except JSONDecodeError:
                msg = f"{self.host} - {debugString} failed to parse json from body. Returned code:{response.status_code} body:{body}"
                logging.error(msg)
                return Result([] if isResponseList else {}, Result.Error(msg))
        else:
            return Result(body, None)
