import asyncio
import os
from distutils.util import strtobool

import pytest
from api.appd.AppDService import AppDService

APPLICATION_ID = int(os.getenv("TEST_CONTROLLER_APPLICATION_ID"))
EUM_APPLICATION_ID = int(os.getenv("TEST_CONTROLLER_EUM_APPLICATION_ID"))
USERNAME = os.getenv("TEST_CONTROLLER_USERNAME")


@pytest.fixture
def event_loop():
    """Overwrite `pytest_asyncio` eventloop to fix Windows issue."""
    if os.name == "nt":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    loop = asyncio.get_event_loop_policy().new_event_loop()
    yield loop
    loop.close()


@pytest.fixture
async def controller():
    if os.name == "nt":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

    host = os.getenv("TEST_CONTROLLER_HOST")
    port = int(os.getenv("TEST_CONTROLLER_PORT"))
    ssl = strtobool(os.getenv("TEST_CONTROLLER_SSL"))
    account = os.getenv("TEST_CONTROLLER_ACCOUNT")
    username = os.getenv("TEST_CONTROLLER_USERNAME")
    pwd = os.getenv("TEST_CONTROLLER_PASSWORD")

    controller = AppDService(
        host=host,
        port=port,
        ssl=ssl,
        account=account,
        username=username,
        pwd=pwd,
    )
    yield controller


@pytest.mark.asyncio
async def testLogin(controller):
    assert (await controller.loginToController()).error is None
    await controller.close()


@pytest.mark.asyncio
async def testGetApmApplications(controller):
    applications = await controller.getApmApplications()
    assert applications.error is None

    application = next(
        (application for application in applications.data if application["id"] == APPLICATION_ID),
        None,
    )
    assert application is not None

    await controller.close()


@pytest.mark.asyncio
async def testGetBtMatchRules(controller):
    assert (await controller.loginToController()).error is None
    btMatchRules = await controller.getBtMatchRules(APPLICATION_ID)

    assert btMatchRules.error is None
    assert "ruleScopeSummaryMappings" in btMatchRules.data
    assert len(btMatchRules.data) > 0

    for rule in btMatchRules.data["ruleScopeSummaryMappings"]:
        assert "rule" in rule
        assert "enabled" in rule["rule"]
        assert "summary" in rule["rule"]
        assert "name" in rule["rule"]["summary"]

    await controller.close()


@pytest.mark.asyncio
async def testGetConfigurations(controller):
    assert (await controller.loginToController()).error is None
    configurations = await controller.getConfigurations()

    assert configurations.error is None
    assert len(configurations.data) > 0

    for configuration in configurations.data:
        assert "name" in configuration
        assert "value" in configuration

    backendLimitProp = next(
        iter([configuration for configuration in configurations.data if configuration["name"] == "backend.registration.limit"]),
        None,
    )
    assert backendLimitProp is not None
    assert "value" in backendLimitProp
    assert int(backendLimitProp["value"]) >= 0

    serviceEndpointLimitProp = next(
        iter([configuration for configuration in configurations.data if configuration["name"] == "sep.ADD.registration.limit"]),
        None,
    )
    assert serviceEndpointLimitProp is not None
    assert "value" in serviceEndpointLimitProp
    assert int(serviceEndpointLimitProp["value"]) >= 0

    await controller.close()


@pytest.mark.asyncio
async def testGetAllCustomExitPoints(controller):
    assert (await controller.loginToController()).error is None
    customExitPoints = await controller.getAllCustomExitPoints(APPLICATION_ID)

    assert customExitPoints.error is None

    for customExitPoint in customExitPoints.data:
        assert "name" in customExitPoint
        assert "agentType" in customExitPoint

    assert (
        next(
            customExitPoint
            for customExitPoint in customExitPoints.data
            if customExitPoint["name"] == "FOO" and customExitPoint["agentType"] == "APP_AGENT"
        ),
        None,
    ) is not None

    await controller.close()


@pytest.mark.asyncio
async def testGetBackendDiscoveryConfigs(controller):
    assert (await controller.loginToController()).error is None
    backendDiscoveryConfigs = await controller.getBackendDiscoveryConfigs(APPLICATION_ID)

    assert backendDiscoveryConfigs.error is None

    for config in backendDiscoveryConfigs.data:
        assert "version" in config

    numberOfModifiedDefaultBackendDiscoveryConfigs = len([config for config in backendDiscoveryConfigs.data if config["version"] != 0])
    assert numberOfModifiedDefaultBackendDiscoveryConfigs != 0

    await controller.close()


@pytest.mark.asyncio
async def testGetDevModeConfig(controller):
    assert (await controller.loginToController()).error is None
    devModeConfig = await controller.getDevModeConfig(APPLICATION_ID)

    assert devModeConfig.error is None

    for config in devModeConfig.data:
        assert "children" in config
        for child in config["children"]:
            assert "enabled" in child

    await controller.close()


@pytest.mark.asyncio
async def testGetInstrumentationLevel(controller):
    assert (await controller.loginToController()).error is None
    instrumentationLevel = await controller.getInstrumentationLevel(APPLICATION_ID)

    assert instrumentationLevel.error is None
    assert instrumentationLevel.data == "DEVELOPMENT" or instrumentationLevel.data == "PRODUCTION"

    await controller.close()


@pytest.mark.asyncio
async def testGetAllNodePropertiesForCustomizedComponents(controller):
    assert (await controller.loginToController()).error is None
    allNodeProperties = await controller.getAllNodePropertiesForCustomizedComponents(APPLICATION_ID)

    assert allNodeProperties.error is None

    for nodeProperties in allNodeProperties.data:
        assert nodeProperties.error is None
        assert "properties" in nodeProperties.data
        for nodeProperty in nodeProperties.data["properties"]:
            assert "stringValue" in nodeProperty
            assert "definition" in nodeProperty
            assert "name" in nodeProperty["definition"]

    await controller.close()


@pytest.mark.asyncio
async def testGetApplicationConfiguration(controller):
    assert (await controller.loginToController()).error is None
    applicationConfiguration = await controller.getApplicationConfiguration(APPLICATION_ID)

    assert applicationConfiguration.error is None

    callGraphConfigurations = {config: value for config, value in applicationConfiguration.data.items() if config == "callGraphConfiguration"}
    assert len(callGraphConfigurations) > 0

    for config in callGraphConfigurations.values():
        assert "hotspotsEnabled" in config
        assert type(config["hotspotsEnabled"]) == bool
        assert "rawSQL" in config
        assert type(config["rawSQL"]) == bool

    await controller.close()


@pytest.mark.asyncio
async def testGetServiceEndpointCustomMatchRules(controller):
    assert (await controller.loginToController()).error is None
    serviceEndpointMatchRules = await controller.getServiceEndpointMatchRules(APPLICATION_ID)

    assert serviceEndpointMatchRules.error is None

    serviceEndpointCustomMatchRules = serviceEndpointMatchRules.data[0]
    serviceEndpointDefaultMatchRules = serviceEndpointMatchRules.data[1]

    for tier in serviceEndpointCustomMatchRules:
        assert tier.error is None
        for definition in tier.data:
            assert "name" in definition
            assert "version" in definition
            assert "agentType" in definition

    for defaultRule in serviceEndpointDefaultMatchRules:
        assert defaultRule.error is None
        for definition in defaultRule.data:
            assert "enabled" in definition
            assert "name" in definition
            assert "version" in definition
            assert "agentType" in definition

    await controller.close()


@pytest.mark.asyncio
async def testGetAppLevelBTConfig(controller):
    assert (await controller.loginToController()).error is None
    appLevelBTConfig = await controller.getAppLevelBTConfig(APPLICATION_ID)

    assert appLevelBTConfig.error is None

    assert "isBtLockDownEnabled" in appLevelBTConfig.data
    assert "isBtAutoCleanupEnabled" in appLevelBTConfig.data
    assert "btAutoCleanupTimeFrame" in appLevelBTConfig.data
    assert "btAutoCleanupCallCountThreshold" in appLevelBTConfig.data

    await controller.close()


@pytest.mark.asyncio
async def testGetMetricData(controller):
    assert (await controller.loginToController()).error is None

    metricData = await controller.getMetricData(
        applicationID=APPLICATION_ID,
        metric_path="Business Transaction Performance|Business Transactions|*|*|Calls per Minute",
        rollup=True,
        time_range_type="BEFORE_NOW",
        duration_in_mins="1440",
    )

    assert metricData.error is None

    assert len(metricData.data) > 0
    for metric in metricData.data:
        assert "metricId" in metric
        assert "metricName" in metric
        assert "metricPath" in metric
        assert "frequency" in metric
        assert "metricValues" in metric
        for value in metric["metricValues"]:
            assert "startTimeInMillis" in value
            assert "occurrences" in value
            assert "current" in value
            assert "min" in value
            assert "max" in value
            assert "useRange" in value
            assert "count" in value
            assert "sum" in value
            assert "value" in value
            assert "standardDeviation" in value

    await controller.close()


@pytest.mark.asyncio
async def testGetEventCountsLastDay(controller):
    assert (await controller.loginToController()).error is None

    eventCounts = await controller.getEventCounts(
        applicationID=APPLICATION_ID,
        entityType="APPLICATION",
        entityID=APPLICATION_ID,
    )

    assert eventCounts.error is None

    assert "policyViolationEventCounts" in eventCounts.data
    assert "totalPolicyViolations" in eventCounts.data["policyViolationEventCounts"]
    assert "warning" in eventCounts.data["policyViolationEventCounts"]["totalPolicyViolations"]
    assert "critical" in eventCounts.data["policyViolationEventCounts"]["totalPolicyViolations"]

    assert eventCounts.data["policyViolationEventCounts"]["totalPolicyViolations"]["warning"] >= 0
    assert eventCounts.data["policyViolationEventCounts"]["totalPolicyViolations"]["critical"] >= 0

    await controller.close()


@pytest.mark.asyncio
async def testGetHealthRules(controller):
    assert (await controller.loginToController()).error is None

    healthRules = await controller.getHealthRules(APPLICATION_ID)

    assert healthRules.error is None

    for healthRule in healthRules.data:
        assert healthRule.error is None

        assert "id" in healthRule.data
        assert "name" in healthRule.data
        assert "enabled" in healthRule.data
        assert "useDataFromLastNMinutes" in healthRule.data
        assert "waitTimeAfterViolation" in healthRule.data
        assert "scheduleName" in healthRule.data
        assert "affects" in healthRule.data
        assert "evalCriterias" in healthRule.data

    await controller.close()


@pytest.mark.asyncio
async def testGetPolicies(controller):
    assert (await controller.loginToController()).error is None

    policies = await controller.getPolicies(APPLICATION_ID)

    assert policies.error is None
    assert len(policies.data) > 0

    for policy in policies.data:
        assert "id" in policy
        assert "name" in policy
        assert "enabled" in policy
        assert "actions" in policy
        assert "events" in policy
        assert "selectedEntityType" in policy

    await controller.close()


@pytest.mark.asyncio
async def testGetPolicies(controller):
    assert (await controller.loginToController()).error is None

    policies = await controller.getPolicies(APPLICATION_ID)

    assert policies.error is None
    assert len(policies.data) > 0

    for policy in policies.data:
        assert "id" in policy
        assert "name" in policy
        assert "enabled" in policy
        assert "actions" in policy
        assert "events" in policy
        assert "selectedEntityType" in policy

    await controller.close()


@pytest.mark.asyncio
async def testGetDataCollectorUsage(controller):
    assert (await controller.loginToController()).error is None

    dataCollectorUsage = await controller.getDataCollectorUsage(APPLICATION_ID)

    assert dataCollectorUsage.error is None

    assert "allDataCollectors" in dataCollectorUsage.data
    assert "dataCollectorsPresentInSnapshots" in dataCollectorUsage.data
    assert "allDataCollectors" in dataCollectorUsage.data

    assert ("HTTP Parameter", "foo", True) in dataCollectorUsage.data["allDataCollectors"]
    assert (
        "Business Data",
        "in_snapshot_not_analytics",
        False,
    ) in dataCollectorUsage.data["allDataCollectors"]
    assert (
        "Business Data",
        "in_snapshot_and_analytics",
        True,
    ) in dataCollectorUsage.data["allDataCollectors"]

    assert not ("HTTP Parameter", "foo", True) in dataCollectorUsage.data["dataCollectorsPresentInSnapshots"]
    assert (
        "Business Data",
        "in_snapshot_not_analytics",
        False,
    ) in dataCollectorUsage.data["dataCollectorsPresentInSnapshots"]
    assert (
        "Business Data",
        "in_snapshot_and_analytics",
        True,
    ) in dataCollectorUsage.data["dataCollectorsPresentInSnapshots"]

    assert not ("HTTP Parameter", "foo", True) in dataCollectorUsage.data["dataCollectorsPresentInAnalytics"]
    assert not ("Business Data", "in_snapshot_not_analytics", False) in dataCollectorUsage.data["dataCollectorsPresentInAnalytics"]
    assert (
        "Business Data",
        "in_snapshot_and_analytics",
        True,
    ) in dataCollectorUsage.data["dataCollectorsPresentInAnalytics"]

    await controller.close()


@pytest.mark.asyncio
async def testGetAnalyticsEnabledStatusForAllApplications(controller):
    assert (await controller.loginToController()).error is None

    analyticsEnabledStatusList = await controller.getAnalyticsEnabledStatusForAllApplications()

    assert analyticsEnabledStatusList.error is None
    assert len(analyticsEnabledStatusList.data) > 0

    for status in analyticsEnabledStatusList.data:
        assert "applicationName" in status
        assert "applicationId" in status
        assert "enabled" in status

    await controller.close()


@pytest.mark.asyncio
async def testGetDashboards(controller):
    assert (await controller.loginToController()).error is None

    dashboards = await controller.getDashboards()

    assert dashboards.error is None
    assert len(dashboards.data) > 0
    for dashboard in dashboards.data:
        assert "name" in dashboard
        assert "widgetTemplates" in dashboard
        assert "createdBy" in dashboard
        assert "createdOn" in dashboard
        assert "modifiedOn" in dashboard

    await controller.close()


@pytest.mark.asyncio
async def testGetUserPermissions(controller):
    assert (await controller.loginToController()).error is None

    userPermissions = await controller.getUserPermissions(USERNAME)

    assert userPermissions.error is None
    assert len(userPermissions.data) > 0

    assert "name" in userPermissions.data
    assert "id" in userPermissions.data
    assert "version" in userPermissions.data

    assert "roles" in userPermissions.data
    for role in userPermissions.data["roles"]:
        assert "name" in role
        assert "id" in role
        assert "version" in role

    adminRole = next(
        (role for role in userPermissions.data["roles"] if role["name"] == "super-admin"),
        None,
    )
    assert adminRole is not None

    await controller.close()


@pytest.mark.asyncio
async def testGetCustomMetrics(controller):
    assert (await controller.loginToController()).error is None

    customMetrics = await controller.getCustomMetrics(APPLICATION_ID, "api-services")

    assert customMetrics.error is None

    await controller.close()


@pytest.mark.asyncio
async def testGetAccountUsageSummary(controller):
    assert (await controller.loginToController()).error is None

    accountUsageSummary = await controller.getAccountUsageSummary()

    assert accountUsageSummary.error is None
    assert len(accountUsageSummary.data) > 0
    for licenseType, licenseData in accountUsageSummary.data.items():
        assert "LicenseProperties" in licenseType
        if licenseData is not None:
            assert "isLicensed" in licenseData
            assert "peakUsage" in licenseData
            assert "numOfProvisionedLicense" in licenseData
            assert "expirationDate" in licenseData

    await controller.close()


@pytest.mark.asyncio
async def testGetAppServerAgents(controller):
    assert (await controller.loginToController()).error is None

    agents = await controller.getAppServerAgents()

    assert agents.error is None
    assert len(agents.data) > 0
    for agent in agents.data:
        assert "hostName" in agent
        assert "agentVersion" in agent
        assert "nodeName" in agent
        assert "componentName" in agent
        assert "applicationName" in agent
        assert "applicationId" in agent
        assert "agentId" in agent
        assert "disabled" in agent
        assert "allDisabled" in agent
        assert "registeredNode" in agent
        assert "applicationComponentNodeId" in agent
        assert "machineId" in agent
        assert "type" in agent

    await controller.close()


@pytest.mark.asyncio
async def testGetMachineAgents(controller):
    assert (await controller.loginToController()).error is None

    agents = await controller.getMachineAgents()

    assert agents.error is None
    assert len(agents.data) > 0
    for agent in agents.data:
        assert "hostName" in agent
        assert "agentVersion" in agent
        assert "applicationNames" in agent
        assert "applicationIds" in agent
        assert "machineId" in agent
        assert "enabled" in agent

    await controller.close()


# @pytest.mark.asyncio
# async def testGetDBAgents(controller):
#     assert (await controller.loginToController()).error is None
#
#     agents = await controller.getDBAgents()
#
#     assert agents.error is None
#     assert len(agents.data) > 0
#     for agent in agents.data:
#         assert "hostName" in agent
#         assert "version" in agent
#         assert "agentName" in agent
#         assert "status" in agent
#         assert "startTime" in agent
#         assert "id" in agent
#
#     await controller.close()


# @pytest.mark.asyncio
# async def testGetAnalyticsAgents(controller):
#     assert (await controller.loginToController()).error is None
#
#     agents = await controller.getAnalyticsAgents()
#
#     assert agents.error is None
#     assert len(agents.data) > 0
#     for agent in agents.data:
#         assert "id" in agent
#         assert "enabled" in agent
#         assert "name" in agent
#         assert "hostName" in agent
#         assert "agentRuntime" in agent
#         assert "installDir" in agent
#         assert "installTimestamp" in agent
#         assert "lastStartTimestamp" in agent
#         assert "lastConnectionTimestamp" in agent
#         assert "majorVersion" in agent
#         assert "minorVersion" in agent
#         assert "pointRelease" in agent
#         assert "agentPointRelease" in agent
#         assert "agentTags" in agent
#         assert "bizTxnsHealthy" in agent
#         assert "logsHealthy" in agent
#
#     await controller.close()


@pytest.mark.asyncio
async def testGetCustomMetrics(controller):
    assert (await controller.loginToController()).error is None

    customMetrics = await controller.getCustomMetrics(APPLICATION_ID, "machine-agent")

    assert customMetrics.error is None
    assert len(customMetrics.data) > 0
    for customMetric in customMetrics.data:
        assert "name" in customMetric
        assert "children" in customMetric
        assert "metricTreeRootType" in customMetric
        assert "metricId" in customMetric
        assert "type" in customMetric
        assert "entityId" in customMetric
        assert "metricPath" in customMetric
        assert "iconPath" in customMetric
        assert "hasChildren" in customMetric

    await controller.close()


@pytest.mark.asyncio
async def testGetEumApplications(controller):
    assert (await controller.loginToController()).error is None

    eumApps = await controller.getEumApplications()

    assert eumApps.error is None
    assert len(eumApps.data) > 0
    for eumApp in eumApps.data:
        assert "name" in eumApp
        assert "appKey" in eumApp
        assert "id" in eumApp
        assert "monitoringEnabled" in eumApp
        assert "synTotalJobs" in eumApp
        assert "metrics" in eumApp

        folders = [
            "syntheticEndUserResponseTime",
            "syntheticPageRequestsPerMin",
            "pageRequestsPerMin",
            "javascriptErrorsPerMin",
            "endUserResponseTime",
            "errorRate",
            "syntheticAvailability",
        ]
        metrics = ["metricId", "name", "value", "sum", "count", "graphData"]

        for folder in folders:
            assert folder in eumApp["metrics"]
            for metric in metrics:
                assert metric in eumApp["metrics"][folder]

    await controller.close()


@pytest.mark.asyncio
async def testGetEumPageListViewData(controller):
    assert (await controller.loginToController()).error is None

    eumPageListViewData = await controller.getEumPageListViewData(EUM_APPLICATION_ID)

    assert eumPageListViewData.error is None
    assert len(eumPageListViewData.data) > 0
    assert "application" in eumPageListViewData.data
    assert "pageIFrameLimit" in eumPageListViewData.data
    assert "ajaxLimit" in eumPageListViewData.data
    assert "id" in eumPageListViewData.data["application"]
    assert "name" in eumPageListViewData.data["application"]

    await controller.close()


@pytest.mark.asyncio
async def testGetPagesAndFramesConfig(controller):
    assert (await controller.loginToController()).error is None

    pagesAndFramesConfig = await controller.getPagesAndFramesConfig(EUM_APPLICATION_ID)

    assert pagesAndFramesConfig.error is None

    assert "customNamingIncludeRules" in pagesAndFramesConfig.data
    assert len(pagesAndFramesConfig.data["customNamingIncludeRules"]) > 0
    for config in pagesAndFramesConfig.data["customNamingIncludeRules"]:
        assert "name" in config
        assert "matchOnMobileApplicationName" in config
        assert "enabled" in config
        assert "priority" in config
        assert "matchOnURL" in config
        assert "isDefault" in config
        assert "pageNamingConfig" in config

    assert "customNamingExcludeRules" in pagesAndFramesConfig.data
    assert len(pagesAndFramesConfig.data["customNamingExcludeRules"]) > 0
    for config in pagesAndFramesConfig.data["customNamingExcludeRules"]:
        assert "name" in config
        assert "matchOnMobileApplicationName" in config
        assert "enabled" in config
        assert "priority" in config
        assert "matchOnURL" in config
        assert "matchOnIpAddressMasks" in config
        assert "matchOnUserAgentType" in config
        assert "matchOnUserAgent" in config

    await controller.close()


@pytest.mark.asyncio
async def testGetAJAXConfig(controller):
    assert (await controller.loginToController()).error is None

    ajaxConfig = await controller.getAJAXConfig(EUM_APPLICATION_ID)

    assert ajaxConfig.error is None

    assert "customNamingIncludeRules" in ajaxConfig.data
    assert len(ajaxConfig.data["customNamingIncludeRules"]) > 0
    for config in ajaxConfig.data["customNamingIncludeRules"]:
        assert "name" in config
        assert "matchOnMobileApplicationName" in config
        assert "enabled" in config
        assert "priority" in config
        assert "matchOnURL" in config
        assert "isDefault" in config
        assert "pageNamingConfig" in config

    assert "customNamingExcludeRules" in ajaxConfig.data
    assert len(ajaxConfig.data["customNamingExcludeRules"]) > 0
    for config in ajaxConfig.data["customNamingExcludeRules"]:
        assert "name" in config
        assert "matchOnMobileApplicationName" in config
        assert "enabled" in config
        assert "priority" in config
        assert "matchOnURL" in config
        assert "matchOnIpAddressMasks" in config
        assert "matchOnUserAgentType" in config
        assert "matchOnUserAgent" in config

    assert "eventServiceIncludeRules" in ajaxConfig.data
    assert len(ajaxConfig.data["eventServiceIncludeRules"]) > 0
    for config in ajaxConfig.data["eventServiceIncludeRules"]:
        assert "name" in config
        assert "matchOnMobileApplicationName" in config
        assert "enabled" in config
        assert "priority" in config
        assert "matchOnURL" in config
        assert "matchBy" in config
        assert "samplePercentage" in config

    assert "eventServiceExcludeRules" in ajaxConfig.data
    assert len(ajaxConfig.data["eventServiceExcludeRules"]) > 0
    for config in ajaxConfig.data["eventServiceExcludeRules"]:
        assert "name" in config
        assert "matchOnMobileApplicationName" in config
        assert "enabled" in config
        assert "priority" in config
        assert "matchOnURL" in config
        assert "matchBy" in config

    await controller.close()


@pytest.mark.asyncio
async def testGetVirtualPagesConfig(controller):
    assert (await controller.loginToController()).error is None

    virtualPagesConfig = await controller.getVirtualPagesConfig(EUM_APPLICATION_ID)

    assert virtualPagesConfig.error is None

    assert "customNamingIncludeRules" in virtualPagesConfig.data
    assert len(virtualPagesConfig.data["customNamingIncludeRules"]) > 0
    for config in virtualPagesConfig.data["customNamingIncludeRules"]:
        assert "name" in config
        assert "matchOnMobileApplicationName" in config
        assert "enabled" in config
        assert "priority" in config
        assert "matchOnURL" in config
        assert "isDefault" in config
        assert "pageNamingConfig" in config

    assert "customNamingExcludeRules" in virtualPagesConfig.data
    assert len(virtualPagesConfig.data["customNamingExcludeRules"]) > 0
    for config in virtualPagesConfig.data["customNamingExcludeRules"]:
        assert "name" in config
        assert "matchOnMobileApplicationName" in config
        assert "enabled" in config
        assert "priority" in config
        assert "matchOnURL" in config
        assert "matchOnIpAddressMasks" in config
        assert "matchOnUserAgentType" in config
        assert "matchOnUserAgent" in config

    await controller.close()
