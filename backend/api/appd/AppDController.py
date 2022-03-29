from uplink import (
    Consumer,
    get,
    params,
    error_handler,
    post,
    Path,
    Body,
    Query,
    headers,
)


class ApiError(Exception):
    pass


def raise_api_error(exc_type, exc_val, exc_tb):
    raise ApiError(exc_val)


@error_handler(raise_api_error)
class AppdController(Consumer):
    """Minimal python client for the AppDynamics API"""

    jsessionid: str = None
    xcsrftoken: str = None

    @params({"action": "login"})
    @get("/controller/auth")
    def login(self):
        """Verifies Login Success"""

    @params({"output": "json"})
    @get("/controller/rest/applications")
    def getApmApplications(self):
        """Retrieves Applications"""

    @params({"output": "json"})
    @get("/controller/rest/applications/{applicationID}/business-transactions")
    def getBTs(self, applicationID: Path):
        """Retrieves Applications"""

    @params({"output": "json"})
    @get("/controller/rest/applications/{applicationID}/nodes")
    def getNodes(self, applicationID: Path):
        """Retrieves Nodes"""

    @params({"output": "json"})
    @get("/controller/rest/applications/{applicationID}/nodes/{nodeID}")
    def getNode(self, applicationID: Path, nodeID: Path):
        """Retrieves an Individual Node"""

    @params({"output": "json"})
    @get("/controller/rest/applications/{applicationID}/tiers")
    def getTiers(self, applicationID: Path):
        """Retrieves Tiers"""

    @params({"output": "json"})
    @get("/controller/restui/transactionConfigProto/getRules/{applicationID}")
    def getBtMatchRules(self, applicationID: Path):
        """Retrieves Business Transaction Match Rules"""

    @params({"output": "json"})
    @get("/controller/restui/transactionConfig/getAppLevelBTConfig/{applicationID}")
    def getAppLevelBTConfig(self, applicationID: Path):
        """Retrieves Application Level Business Transaction Configurations"""

    @params({"output": "json"})
    @get("/controller/rest/applications/{applicationID}/backends")
    def getBackends(self, applicationID: Path):
        """Retrieves Backends"""

    @params({"output": "json"})
    @get("/controller/rest/configuration")
    def getConfigurations(self):
        """Retrieves Controller Configurations"""

    @params({"output": "json"})
    @post("/controller/restui/customExitPoint/getAllCustomExitPoints")
    def getAllCustomExitPoints(self, application: Body):
        """Retrieves Custom Edit Point Configurations"""

    @params({"output": "json"})
    @post("/controller/restui/backendConfig/getBackendDiscoveryConfigs")
    def getBackendDiscoveryConfigs(self, body: Body):
        """Retrieves Controller Configurations"""

    @params({"output": "json"})
    @get("/controller/restui/applicationManagerUiBean/getDevModeConfig/{applicationID}")
    def getDevModeConfig(self, applicationID: Path):
        """Retrieves Developer Mode Configuration"""

    @params({"output": "json"})
    @get("/controller/restui/applicationManagerUiBean/instrumentationLevel/{applicationID}")
    def getInstrumentationLevel(self, applicationID: Path):
        """Retrieves Instrumentation Level"""

    @params({"output": "json"})
    @get("/controller/restui/agentManager/getAllApplicationComponentsWithNodes/{applicationID}")
    def getAllApplicationComponentsWithNodes(self, applicationID: Path):
        """Retrieves Node Configurations"""

    @params({"output": "json"})
    @headers(
        {
            "Accept": "application/json, text/plain, */*",
        }
    )
    @post("/controller/restui/agentManager/getAgentConfiguration")
    def getAgentConfiguration(self, body: Body):
        """Retrieves Agent Configurations"""

    @params({"output": "json"})
    @get("/controller/restui/applicationManagerUiBean/applicationConfiguration/{applicationID}")
    def getApplicationConfiguration(self, applicationID: Path):
        """Retrieves Application Configuration"""

    @params({"output": "json"})
    @get("/controller/restui/components/application/{applicationID}/components")
    def getApplicationComponents(self, applicationID: Path):
        """Retrieves Application Components for Later  to get getServiceEndpointCustomMatchRules"""

    @params({"output": "json"})
    @post("/controller/restui/serviceEndpoint/getAll")
    def getServiceEndpointCustomMatchRules(self, body: Body):
        """Retrieves Service Endpoint Custom Match Rules for an individual Application Tier"""

    @params({"output": "json"})
    @post("/controller/restui/serviceEndpoint/getServiceEndpointMatchConfigs")
    def getServiceEndpointDefaultMatchRules(self, body: Body):
        """Retrieves Service Endpoint Custom Match Rules for an individual Application Tier"""

    @params({"output": "json"})
    @get("/controller/restui/events/eventCounts?timeRangeString=last_1_day.BEFORE_NOW.-1.-1.1440")
    def getEventCountsLastDay(
        self,
        applicationID: Query("applicationId"),
        entityType: Query("entityType"),
        entityID: Query("entityId"),
    ):
        """Retrieves Event Counts from the past day"""

    @params({"output": "json"})
    @post("/controller/restui/metricBrowser/async/metric-tree/root")
    def getMetricTree(self, body: Body):
        """Retrieves Metrics"""

    @params({"output": "json"})
    @get("/controller/rest/applications/{applicationID}/metric-data")
    def getMetricData(
        self,
        applicationID: Path,
        metric_path: Query("metric-path"),
        rollup: Query("rollup"),
        time_range_type: Query("time-range-type"),
        duration_in_mins: Query("duration-in-mins"),
        start_time: Query("start-time"),
        end_time: Query("end-time"),
    ):
        """Retrieves Metrics"""

    @params({"output": "json"})
    @get("/controller/rest/applications/{applicationID}/events")
    def getApplicationEvents(
        self,
        applicationID: Path,
        event_types: Query("event-types"),
        severities: Query("severities"),
        time_range_type: Query("time-range-type"),
        duration_in_mins: Query("duration-in-mins"),
        start_time: Query("start-time"),
        end_time: Query("end-time"),
    ):
        """Retrieves Events"""

    @params({"output": "json"})
    @get("/controller/alerting/rest/v1/applications/{applicationID}/health-rules")
    def getHealthRules(self, applicationID: Path):
        """Retrieves Health Rules"""

    @params({"output": "json"})
    @get("/controller/alerting/rest/v1/applications/{applicationID}/health-rules/{healthRuleID}")
    def getHealthRule(self, applicationID: Path, healthRuleID: Path):
        """Retrieves Specific Health Rule"""

    @params({"output": "json"})
    @get("/controller/alerting/rest/v1/applications/{applicationID}/policies")
    def getPolicies(self, applicationID: Path):
        """Retrieves Policies"""

    @params({"output": "json"})
    @get("/controller/restui/MidcUiService/getAllDataGathererConfigs/{applicationID}")
    def getDataCollectors(self, applicationID: Path):
        """Retrieves Data Collectors"""

    @params({"output": "json"})
    @post("/controller/restui/snapshot/snapshotListDataWithFilterHandle")
    def getSnapshotsWithDataCollector(self, body: Body):
        """Retrieves Snapshots"""

    @params({"output": "json"})
    @get("/controller/restui/analyticsConfigTxnAnalyticsUiService/getAllVisibleAppsWithAnalyticsInfo")
    def getAnalyticsEnabledStatusForAllApplications(self):
        """Retrieves Analytics Enabled Status for app Applications"""

    @params({"output": "json"})
    @get("/controller/restui/dashboards/getAllDashboardsByType/false")
    def getAllDashboardsMetadata(self):
        """Retrieves all Dashboards"""

    @params({"output": "json"})
    @get("/controller/CustomDashboardImportExportServlet")
    def getDashboard(self, dashboardId: Query("dashboardId")):
        """Retrieves a single Dashboard"""

    @params({"output": "json"})
    @get("/controller/restui/userAdministrationUiService/users")
    def getUsers(self):
        """Retrieves list of Users"""

    @params({"output": "json"})
    @get("/controller/restui/userAdministrationUiService/users/{userID}")
    def getUser(self, userID: Path):
        """Retrieves permission set of a given user"""

    @params({"output": "json"})
    @post("/controller/restui/licenseRule/getAllLicenseModuleProperties")
    def getAccountUsageSummary(self, body: Body):
        """Retrieves license usage summary"""

    @params({"output": "json"})
    @post("/controller/restui/agents/list/appserver")
    def getAppServerAgents(self, body: Body):
        """Retrieves app server agent summary list"""

    @params({"output": "json"})
    @post("/controller/restui/agents/list/machine")
    def getMachineAgents(self, body: Body):
        """Retrieves machine agent summary list"""

    @params({"output": "json"})
    @post("/controller/restui/agents/list/appserver/ids")
    def getAppServerAgentsIds(self, body: Body):
        """Retrieves app server agent summary list"""

    @params({"output": "json"})
    @post("/controller/restui/agents/list/machine/ids")
    def getMachineAgentsIds(self, body: Body):
        """Retrieves machine agent summary list"""

    @params({"output": "json"})
    @get("/controller/restui/agent/setting/getDBAgents")
    def getDBAgents(self):
        """Retrieves db agent summary list"""

    @params({"output": "json"})
    @get("/controller/restui/analytics/agents/agentsStatus")
    def getAnalyticsAgents(self):
        """Retrieves analytics agent summary list"""

    @params({"output": "json", "time-range": "last_1_day.BEFORE_NOW.-1.-1.1440"})
    @get("/controller/restui/eumApplications/getAllEumApplicationsData")
    def getEumApplications(self):
        """Retrieves all Eum Applications"""

    @params({"output": "json"})
    @post("/controller/restui/pageList/getEumPageListViewData")
    def getEumPageListViewData(self, body: Body):
        """Retrieves Eum Page List View Data"""

    @params({"output": "json"})
    @post("/controller/restui/web/pagelist")
    def getEumNetworkRequestList(self, body: Body):
        """Retrieves Eum Network Request List"""

    @params({"output": "json"})
    @get("/controller/restui/browserRUMConfig/getPagesAndFramesConfig/{applicationId}")
    def getPagesAndFramesConfig(self, applicationId: Path):
        """Retrieves pages and frames config"""

    @params({"output": "json"})
    @get("/controller/restui/browserRUMConfig/getAJAXConfig/{applicationId}")
    def getAJAXConfig(self, applicationId: Path):
        """Retrieves AJAX config"""

    @params({"output": "json"})
    @get("/controller/restui/browserRUMConfig/getVirtualPagesConfig/{applicationId}")
    def getVirtualPagesConfig(self, applicationId: Path):
        """Retrieves virtual pages config"""

    @params({"output": "json", "time-range": "last_1_day.BEFORE_NOW.-1.-1.1440"})
    @get("/controller/restui/eumApplications/getAllMobileApplicationsData")
    def getMRUMApplications(self):
        """Retrieves all Mrum Applications"""

    @params({"output": "json"})
    @get("/controller/restui/mobileRUMConfig/networkRequestsConfig/{applicationId}")
    def getMRUMNetworkRequestConfig(self, applicationId: Path):
        """Retrieves Mrum network requests config"""

    @params({"output": "json"})
    @get("/controller/restui/mobileRequestListUiService/getNetworkRequestLimit/{applicationId}")
    def getNetworkRequestLimit(self, applicationId: Path):
        """Retrieves network request limit"""
