import ipaddress
import json
import logging
import re
import time
from json import JSONDecodeError
from math import ceil
from typing import List

import aiohttp
from api.appd.AppDController import AppdController
from api.Result import Result
from uplink import AiohttpClient
from uplink.auth import BasicAuth, MultiAuth, ProxyAuth
from util.asyncio_utils import AsyncioUtils


class AppDService:
    controller: AppdController

    def __init__(
        self,
        host: str,
        port: int,
        ssl: bool,
        account: str,
        username: str,
        pwd: str,
        verifySsl: bool = True,
        useProxy: bool = False,
    ):
        logging.debug(f"{host} - Initializing controller service")
        connection_url = f'{"https" if ssl else "http"}://{host}:{port}'
        auth = BasicAuth(f"{username}@{account}", pwd)
        self.host = host
        self.username = username

        cookie_jar = aiohttp.CookieJar()
        try:
            if ipaddress.ip_address(host):
                logging.warning(f"Configured host {host} is an IP address. Consider using the DNS instead.")
                logging.warning(f"RFC 2109 explicitly forbids cookie accepting from URLs with IP address instead of DNS name.")
                logging.warning(f"Using unsafe Cookie Jar.")
                cookie_jar = aiohttp.CookieJar(unsafe=True)
        except ValueError:
            pass

        connector = aiohttp.TCPConnector(limit=AsyncioUtils.concurrentConnections, verify_ssl=verifySsl)
        self.session = aiohttp.ClientSession(connector=connector, trust_env=useProxy, cookie_jar=cookie_jar)

        self.controller = AppdController(
            base_url=connection_url,
            auth=auth,
            client=AiohttpClient(session=self.session),
        )
        self.totalCallsProcessed = 0

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
            logging.error(f"{self.host} - Controller login failed with {response.status_code}. Check username and password.")
            return Result(
                response,
                Result.Error(f"{self.host} - Controller login failed with {response.status_code}. Check username and password."),
            )
        try:
            jsessionid = re.search("JSESSIONID=(\\w|\\d)*", str(response.headers)).group(0).split("JSESSIONID=")[1]
            self.controller.jsessionid = jsessionid
        except AttributeError:
            logging.debug(f"{self.host} - Unable to find JSESSIONID in login response. Please verify credentials.")
        try:
            xcsrftoken = re.search("X-CSRF-TOKEN=(\\w|\\d)*", str(response.headers)).group(0).split("X-CSRF-TOKEN=")[1]
            self.controller.xcsrftoken = xcsrftoken
        except AttributeError:
            logging.debug(f"{self.host} - Unable to find X-CSRF-TOKEN in login response. Please verify credentials.")

        if self.controller.jsessionid is None or self.controller.xcsrftoken is None:
            return Result(
                response,
                Result.Error(f"{self.host} - Valid authentication headers not cached from previous login call. Please verify credentials."),
            )

        self.controller.session.headers["X-CSRF-TOKEN"] = self.controller.xcsrftoken
        self.controller.session.headers["Set-Cookie"] = f"JSESSIONID={self.controller.jsessionid};X-CSRF-TOKEN={self.controller.xcsrftoken};"
        self.controller.session.headers["Content-Type"] = "application/json;charset=UTF-8"

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
        response = await self.controller.getApmApplications()
        result = await self.getResultFromResponse(response, debugString)
        # apparently it's possible to have a null application name, the controller converts the null into "null"
        if result.error is None:
            for application in result.data:
                if application["name"] is None:
                    application["name"] = "null"
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
        return await self.getResultFromResponse(response, debugString, isResponseJSON=False)

    async def getAllNodePropertiesForCustomizedComponents(self, applicationID: int) -> Result:
        debugString = f"Gathering All Application Components With Nodes for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getAllApplicationComponentsWithNodes(applicationID)
        applicationComponentsWithNodes = await self.getResultFromResponse(response, debugString)

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

    async def getAgentConfiguration(self, applicationID: int, agentType: str, entityType: str, entityId: int) -> Result:
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
        response = await self.controller.getApplicationConfiguration(applicationID)
        return await self.getResultFromResponse(response, debugString)

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
        defaultMatchRulesFutures.append(self.controller.getServiceEndpointDefaultMatchRules(body))
        for entity in response.data:
            body = '{"attachedEntity":{"entityType":"APPLICATION_COMPONENT","entityId":{entityId}},"agentType":"{agentType}"}'.replace(
                "{entityId}", str(entity["id"])
            ).replace("{agentType}", str(entity["componentType"]["agentType"]))
            customMatchRulesFutures.append(self.controller.getServiceEndpointCustomMatchRules(body))

            body = '{"agentType":"APP_AGENT","attachedEntity":{"entityId":{entityId},"entityType":"APPLICATION_COMPONENT"}}'.replace(
                "{entityId}", str(entity["id"])
            )
            defaultMatchRulesFutures.append(self.controller.getServiceEndpointDefaultMatchRules(body))

        response = await AsyncioUtils.gatherWithConcurrency(*customMatchRulesFutures)
        customMatchRules = [await self.getResultFromResponse(response, debugString) for response in response]
        response = await AsyncioUtils.gatherWithConcurrency(*defaultMatchRulesFutures)
        defaultMatchRules = [await self.getResultFromResponse(response, debugString) for response in response]

        return Result((customMatchRules, defaultMatchRules), None)

    async def getAppLevelBTConfig(self, applicationID: int) -> Result:
        debugString = f"Gathering Application Business Transaction Configuration Settings for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getAppLevelBTConfig(applicationID)
        return await self.getResultFromResponse(response, debugString)

    async def getCustomMetrics(self, applicationID: int, tierName: str) -> Result:
        debugString = f"Gathering Custom Metrics for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "request": None,
            "applicationId": applicationID,
            "livenessStatus": "ALL",
            "pathData": ["Application Infrastructure Performance", tierName, "Custom Metrics"],
            "timeRangeSpecifier": {
                "type": "BEFORE_NOW",
                "durationInMinutes": 60,
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
        duration_in_mins: str = "",
        start_time: str = "",
        end_time: str = "",
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

    async def getEventCountsLastDay(self, applicationID: int, entityType: str, entityID: int) -> Result:
        debugString = f'Gathering Event Counts for:"{entityType}" {entityID} on application:{applicationID}'
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getEventCountsLastDay(applicationID, entityType, entityID)
        return await self.getResultFromResponse(response, debugString)

    async def getHealthRules(self, applicationID: int) -> Result:
        debugString = f"Gathering Health Rules for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getHealthRules(applicationID)
        healthRules = await self.getResultFromResponse(response, debugString)

        healthRuleDetails = []
        for healthRule in healthRules.data:
            healthRuleDetails.append(self.controller.getHealthRule(applicationID, healthRule["id"]))

        responses = await AsyncioUtils.gatherWithConcurrency(*healthRuleDetails)

        healthRulesData = []
        for response, healthRule in zip(responses, healthRules.data):
            debugString = f"Gathering Health Rule Data for Application:{applicationID} HealthRule:'{healthRule['name']}'"
            healthRulesData.append(await self.getResultFromResponse(response, debugString))

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
        body = (
            '{"firstInChain":false,"maxRows":{maximum_results},"applicationIds":[{applicationID}],"businessTransactionIds":[],"applicationComponentIds":[],"applicationComponentNodeIds":[],"errorIDs":[],"errorOccured":null,"userExperience":[],"executionTimeInMilis":null,"endToEndLatency":null,"url":null,"sessionId":null,"userPrincipalId":null,"dataCollectorFilter":{"collectorType":"{dataCollectorType}","query":{"name":"{dataCollectorName}","value":""}},"archived":null,"guids":[],"diagnosticSnapshot":null,"badRequest":null,"deepDivePolicy":[],"rangeSpecifier":{"type":"BEFORE_NOW","durationInMinutes":1440}}'.replace(
                "{applicationID}", str(applicationID)
            )
            .replace("{dataCollectorType}", str(data_collector_type))
            .replace("{dataCollectorName}", str(data_collector_name))
            .replace("{dataCollectorValue}", str(data_collector_value))
            .replace("{maximum_results}", str(maximum_results))
        )

        response = await self.controller.getSnapshotsWithDataCollector(body)

        return await self.getResultFromResponse(response, debugString)

    async def getDataCollectorUsage(self, applicationID: int) -> Result:
        debugString = f"Gathering Data Collectors for Application:{applicationID}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getDataCollectors(applicationID)

        dataCollectors = await self.getResultFromResponse(response, debugString)
        snapshotEnabledDataCollectors = [dataCollector for dataCollector in dataCollectors.data if dataCollector["enabledForApm"]]

        httpDataCollectors = [dataCollector for dataCollector in snapshotEnabledDataCollectors if dataCollector["type"] == "http"]
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
                dataCollectorFields.append(("Cookie", field, dataCollector["enabledForAnalytics"]))
            for field in dataCollector["sessionKeys"]:
                dataCollectorFields.append(("Session Key", field, dataCollector["enabledForAnalytics"]))
            for field in dataCollector["headers"]:
                dataCollectorFields.append(("HTTP Header", field, dataCollector["enabledForAnalytics"]))

        pojoDataCollectors = [dataCollector for dataCollector in snapshotEnabledDataCollectors if dataCollector["type"] == "pojo"]
        for dataCollector in pojoDataCollectors:
            for methodDataGathererConfig in dataCollector["methodDataGathererConfigs"]:
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
            if (applicationID, dataCollectorField[1], dataCollectorField[0]) not in distinctDataCollectors:
                snapshotsContainingDataCollectorFields.append(
                    self.getSnapshotsWithDataCollector(
                        applicationID=applicationID,
                        data_collector_name=dataCollectorField[1],
                        data_collector_type=dataCollectorField[0],
                    )
                )
            distinctDataCollectors.add((applicationID, dataCollectorField[1], dataCollectorField[0]))
        snapshotResults = await AsyncioUtils.gatherWithConcurrency(*snapshotsContainingDataCollectorFields)

        dataCollectorFieldsWithSnapshots = []
        for collector, snapshotResult in zip(dataCollectorFields, snapshotResults):
            if snapshotResult.error is None and len(snapshotResult.data["requestSegmentDataListItems"]) == 1:
                dataCollectorFieldsWithSnapshots.append(collector)
            # This API does not work for either session keys or headers, as far as I know there is no way to get this info without inspecting ALL snapshots (won't do).
            # The API comes from the Transaction Snapshot filtering UI. No UI option for session keys or headers exists there.
            # For now, let's just assume that any session key or header configured data collector is working... If anyone has a better idea I'm all ears.
            elif collector[0] == "Session Key" or collector[0] == "HTTP Header":
                dataCollectorFieldsWithSnapshots.append(collector)

        result = {
            "allDataCollectors": dataCollectorFields,
            "dataCollectorsPresentInSnapshots": dataCollectorFieldsWithSnapshots,
            "dataCollectorsPresentInAnalytics": [dataCollector for dataCollector in dataCollectorFieldsWithSnapshots if dataCollector[2]],
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
        allDashboardsMetadata = await self.getResultFromResponse(response, debugString)

        dashboards = []
        batch_size = AsyncioUtils.concurrentConnections
        for i in range(0, len(allDashboardsMetadata.data), batch_size):
            dashboardsFutures = []

            logging.debug(f"Batch iteration {int(i / batch_size)} of {ceil(len(allDashboardsMetadata.data) / batch_size)}")
            chunk = allDashboardsMetadata.data[i : i + batch_size]

            for dashboard in chunk:
                dashboardsFutures.append(self.controller.getDashboard(dashboard["id"]))

            response = await AsyncioUtils.gatherWithConcurrency(*dashboardsFutures)
            for dashboard in [await self.getResultFromResponse(response, debugString) for response in response]:
                dashboards.append(dashboard.data)

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
        for dashboardSchema, dashboardOverview in zip(dashboards, allDashboardsMetadata.data):
            if "schemaVersion" in dashboardSchema:
                dashboardSchema["createdBy"] = dashboardOverview["createdBy"]
                dashboardSchema["createdOn"] = dashboardOverview["createdOn"]
                dashboardSchema["modifiedOn"] = dashboardOverview["modifiedOn"]
                returnedDashboards.append(dashboardSchema)

        return Result(returnedDashboards, None)

    async def getUserPermissions(self, username: str) -> Result:
        debugString = f"Gathering Permission set for user: {username}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getUsers()
        users = await self.getResultFromResponse(response, debugString)

        if users.error is not None:
            logging.error(f"{self.host} - Call to Get User Permissions failed. Is user '{self.username}' an Account Owner?")
            return Result(
                response,
                Result.Error(f"{self.host} - Call to Get User Permissions failed. Is user '{self.username}' an Account Owner?"),
            )

        userID = next(user["id"] for user in users.data if user["name"].lower() == username.lower())

        response = await self.controller.getUser(userID)
        return await self.getResultFromResponse(response, debugString)

    async def getAccountUsageSummary(self) -> Result:
        debugString = f"Gathering Account Usage Summary"
        logging.debug(f"{self.host} - {debugString}")
        body = {"type": "BEFORE_NOW", "durationInMinutes": 1440, "endTime": None, "startTime": None, "timeRange": None, "timeRangeAdjusted": False}
        response = await self.controller.getAccountUsageSummary(json.dumps(body))
        return await self.getResultFromResponse(response, debugString)

    async def getAppServerAgents(self) -> Result:
        debugString = f"Gathering App Server Agents Agents"
        logging.debug(f"{self.host} - {debugString}")
        # get current timestamp in milliseconds
        currentTime = int(round(time.time() * 1000))
        # get the last 24 hours in milliseconds
        last24Hours = currentTime - (24 * 60 * 60 * 1000)
        body = {
            "requestFilter": {"queryParams": {"applicationAssociationType": "ALL"}, "filters": []},
            "resultColumns": [],
            "offset": 0,
            "limit": -1,
            "searchFilters": [],
            "columnSorts": [{"column": "HOST_NAME", "direction": "ASC"}],
            "timeRangeStart": last24Hours,
            "timeRangeEnd": currentTime,
        }
        response = await self.controller.getAppServerAgents(json.dumps(body))
        result = await self.getResultFromResponse(response, debugString)
        if result.error is not None:
            return result

        agentIds = [agent["applicationComponentNodeId"] for agent in result.data["data"]]

        debugString = f"Gathering App Server Agents Agents List"
        allAgents = []
        batch_size = 50
        for i in range(0, len(agentIds), batch_size):
            agentFutures = []

            logging.debug(f"Batch iteration {int(i / batch_size)} of {ceil(len(agentIds) / batch_size)}")
            chunk = agentIds[i : i + batch_size]

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
                "timeRangeStart": last24Hours,
                "timeRangeEnd": currentTime,
            }

            response = await self.controller.getAppServerAgentsIds(json.dumps(body))
            result = await self.getResultFromResponse(response, debugString)

            if result.error is None:
                allAgents.extend(result.data["data"])
            else:
                logging.warning(f"{self.host} - Failed to get App Server Agents: {result.error}")

        return Result(allAgents, None)

    async def getMachineAgents(self) -> Result:
        debugString = f"Gathering App Server Agents Agents"
        logging.debug(f"{self.host} - {debugString}")
        # get current timestamp in milliseconds
        currentTime = int(round(time.time() * 1000))
        # get the last 24 hours in milliseconds
        last24Hours = currentTime - (1 * 60 * 60 * 1000)
        body = {
            "requestFilter": {"queryParams": {"applicationAssociationType": "ALL"}, "filters": []},
            "resultColumns": [],
            "offset": 0,
            "limit": -1,
            "searchFilters": [],
            "columnSorts": [{"column": "HOST_NAME", "direction": "ASC"}],
            "timeRangeStart": last24Hours,
            "timeRangeEnd": currentTime,
        }
        response = await self.controller.getMachineAgents(json.dumps(body))
        result = await self.getResultFromResponse(response, debugString)
        if result.error is not None:
            return result

        agentIds = [agent["machineId"] for agent in result.data["data"]]

        debugString = f"Gathering Machine Agents Agents List"
        allAgents = []
        batch_size = 50
        for i in range(0, len(agentIds), batch_size):
            agentFutures = []

            logging.debug(f"Batch iteration {int(i / batch_size)} of {ceil(len(agentIds) / batch_size)}")
            chunk = agentIds[i : i + batch_size]

            body = {
                "requestFilter": chunk,
                "resultColumns": ["AGENT_VERSION", "APPLICATION_NAMES", "ENABLED"],
                "offset": 0,
                "limit": -1,
                "searchFilters": [],
                "columnSorts": [{"column": "HOST_NAME", "direction": "ASC"}],
                "timeRangeStart": last24Hours,
                "timeRangeEnd": currentTime,
            }

            response = await self.controller.getMachineAgentsIds(json.dumps(body))
            result = await self.getResultFromResponse(response, debugString)

            if result.error is None:
                allAgents.extend(result.data["data"])
            else:
                logging.warning(f"{self.host} - Failed to get Machine Agents: {result.error}")

        return Result(allAgents, None)

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

    async def getEumApplications(self) -> Result:
        debugString = f"Gathering EUM Applications"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getEumApplications()
        return await self.getResultFromResponse(response, debugString)

    async def getEumPageListViewData(self, applicationId: int) -> Result:
        debugString = f"Gathering EUM Page List View Data for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        body = {"applicationId": applicationId, "addId": None, "timeRangeString": "last_1_hour|BEFORE_NOW|-1|-1|60", "fetchSyntheticData": False}
        response = await self.controller.getEumPageListViewData(json.dumps(body))
        return await self.getResultFromResponse(response, debugString)

    async def getEumNetworkRequestList(self, applicationId: int) -> Result:
        debugString = f"Gathering EUM Page List View Data for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")

        # get current timestamp in milliseconds
        currentTime = int(round(time.time() * 1000))
        # get the last 24 hours in milliseconds
        last24Hours = currentTime - (1 * 60 * 60 * 1000)

        body = {
            "requestFilter": {"applicationId": applicationId, "fetchSyntheticData": False},
            "resultColumns": ["PAGE_TYPE", "PAGE_NAME", "TOTAL_REQUESTS", "END_USER_RESPONSE_TIME", "VISUALLY_COMPLETE_TIME"],
            "offset": 0,
            "limit": -1,
            "searchFilters": [],
            "columnSorts": [{"column": "TOTAL_REQUESTS", "direction": "DESC"}],
            "timeRangeStart": last24Hours,
            "timeRangeEnd": currentTime,
        }
        response = await self.controller.getEumNetworkRequestList(json.dumps(body))
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

    async def getBrowserSnapshotsWithServerSnapshots(self, applicationId: int) -> Result:
        debugString = f"Gathering Browser Snapshots for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "applicationId": applicationId,
            "timeRangeString": "last_1_hour.BEFORE_NOW.-1.-1.60",
            "filters": {
                "_classType": "BrowserSnapshotFilters",
                "serverSnapshotExists": {"type": "BOOLEAN", "name": "ms_serverSnapshotExists", "value": True},
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
        return await self.getResultFromResponse(response, debugString, isResponseList=False)

    async def getMRUMApplications(self) -> Result:
        debugString = f"Gathering MRUM Applications"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getMRUMApplications()
        return await self.getResultFromResponse(response, debugString)

    async def getMRUMNetworkRequestConfig(self, applicationId: int) -> Result:
        debugString = f"Gathering MRUM Network Request Config for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getMRUMNetworkRequestConfig(applicationId)
        return await self.getResultFromResponse(response, debugString)

    async def getNetworkRequestLimit(self, applicationId: int) -> Result:
        debugString = f"Gathering MRUM Network Request Limit for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getNetworkRequestLimit(applicationId)
        return await self.getResultFromResponse(response, debugString)

    async def getMobileSnapshotsWithServerSnapshots(self, applicationId: int, mobileApplicationId: int, platform: str) -> Result:
        debugString = f"Gathering Mobile Snapshots for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        body = {
            "applicationId": applicationId,
            "timeRangeString": "last_1_hour|BEFORE_NOW|-1|-1|60",
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

    async def getSyntheticBillableTime(self, applicationId: int, scheduleIds: list[str]) -> Result:
        debugString = f"Gathering Synthetic Billable Time for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        # get current timestamp in milliseconds
        currentTime = int(round(time.time() * 1000))
        # get the last 24 hours in milliseconds
        last24Hours = currentTime - (1 * 60 * 60 * 1000)

        body = {
            "scheduleIds": scheduleIds,
            "appId": applicationId,
            "startTime": last24Hours,
            "currentTime": currentTime,
        }
        response = await self.controller.getSyntheticBillableTime(json.dumps(body))
        return await self.getResultFromResponse(response, debugString)

    async def getSyntheticPrivateAgentUtilization(self, applicationId: int, jobsJson: list[dict]) -> Result:
        debugString = f"Gathering Synthetic Private Agent Utilization for Application {applicationId}"
        logging.debug(f"{self.host} - {debugString}")
        response = await self.controller.getSyntheticPrivateAgentUtilization(applicationId, json.dumps(jobsJson))
        return await self.getResultFromResponse(response, debugString)

    async def close(self):
        logging.debug(f"{self.host} - Closing connection")
        await self.session.close()

    async def getResultFromResponse(self, response, debugString, isResponseJSON=True, isResponseList=True) -> Result:
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
            return Result([] if isResponseList else {}, Result.Error(f"{response.status_code}"))
        if isResponseJSON:
            try:
                return Result(json.loads(body), None)
            except JSONDecodeError:
                msg = f"{self.host} - {debugString} failed to parse json from body. Returned code:{response.status_code} body:{body}"
                logging.error(msg)
                return Result([] if isResponseList else {}, Result.Error(msg))
        else:
            return Result(body, None)
