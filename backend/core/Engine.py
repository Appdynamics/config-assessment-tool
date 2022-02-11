import asyncio
import json
import logging
import os
import sys
import time
import traceback
from collections import OrderedDict
from pathlib import Path

from api.appd.AppDService import AppDService
from extractionSteps.bsg.ApmDashboards import ApmDashboards
from extractionSteps.bsg.AppAgents import AppAgents
from extractionSteps.bsg.Backends import Backends
from extractionSteps.bsg.BusinessTransactions import BusinessTransactions
from extractionSteps.bsg.DataCollectors import DataCollectors
from extractionSteps.bsg.ErrorConfiguration import ErrorConfiguration
from extractionSteps.bsg.HealthRulesAndAlerting import HealthRulesAndAlerting
from extractionSteps.bsg.MachineAgents import MachineAgents
from extractionSteps.bsg.OverallAssessment import OverallAssessment
from extractionSteps.bsg.Overhead import Overhead
from extractionSteps.bsg.ServiceEndpoints import ServiceEndpoints
from extractionSteps.general.CustomMetrics import CustomMetrics
from reports.reports.AgentMatrixReport import AgentMatrixReport
from reports.reports.BSGReport import BSGReport
from reports.reports.CustomMetricsReport import CustomMetricsReport
from reports.reports.LicenseReport import LicenseReport
from reports.reports.RawBSGReport import RawDataReport
from util.asyncio_utils import gatherWithConcurrency
from util.stdlib_utils import jsonEncoder


class Engine:
    def __init__(self, jobFileName: str, thresholdsFileName: str):
        logging.info(f'\n{open(f"backend/resources/img/splash.txt").read()}')

        # Validate jobFileName and thresholdFileName
        self.jobFileName = jobFileName
        self.thresholdsFileName = thresholdsFileName
        if not Path(f"input/jobs/{self.jobFileName}.json").exists():
            logging.error(f"Job file input/jobs/{self.jobFileName}.json does not exit. Aborting.")
            sys.exit(1)
        if not Path(f"input/thresholds/{self.thresholdsFileName}.json").exists():
            logging.error(f"Job file input/thresholds/{self.thresholdsFileName}.json does not exit. Aborting.")
            sys.exit(1)
        if not os.path.exists(f"output/{self.jobFileName}"):
            os.makedirs(f"output/{self.jobFileName}")
        self.job = json.loads(open(f"input/jobs/{self.jobFileName}.json").read())
        self.thresholds = json.loads(open(f"input/thresholds/{self.thresholdsFileName}.json").read())

        # Instantiate controllers, jobs, and report lists
        self.controllers = [
            AppDService(
                host=controller["host"],
                port=controller["port"],
                ssl=controller["ssl"],
                account=controller["account"],
                username=controller["username"],
                pwd=controller["pwd"],
                verifySsl=controller.get("verifySsl", True),
                proxyUsername=controller.get("proxyUsername"),
                proxyPassword=controller.get("proxyPassword")
            )
            for controller in self.job
        ]
        self.controllerData = OrderedDict()
        self.bsgSteps = [
            AppAgents(),
            MachineAgents(),
            BusinessTransactions(),
            Backends(),
            Overhead(),
            ServiceEndpoints(),
            ErrorConfiguration(),
            HealthRulesAndAlerting(),
            DataCollectors(),
            ApmDashboards(),
            OverallAssessment(),
        ]
        self.otherSteps = [
            CustomMetrics(),
        ]
        self.reports = [
            BSGReport(),
            RawDataReport(),
            AgentMatrixReport(),
            CustomMetricsReport(),
            LicenseReport(),
        ]

    async def run(self):
        startTime = time.monotonic()

        try:
            self.validateThresholdsFile()
            await self.initControllers()
            await self.process()
            self.finalize()
        except Exception as e:
            # catch exceptions here so we can terminate coroutines before program exit
            logging.error("".join(traceback.TracebackException.from_exception(e).format()))

        await self.abortAndCleanup(
            f"Program completed in {str(round((time.monotonic() - startTime), 2))} seconds",
            error=False,
        )

    async def initControllers(self) -> ([AppDService], str):
        logging.info(f"Validating Controller Login(s) for Job - {self.jobFileName} ")

        loginFutures = [controller.loginToController() for controller in self.controllers]
        loginResults = await gatherWithConcurrency(*loginFutures)
        if any(login.error is not None for login in loginResults):
            await self.abortAndCleanup(f"Unable to connect to one or more controllers. Aborting.")

        userPermissionFutures = [controller.getUserPermissions(controller.username) for controller in self.controllers]
        userPermissionsResults = await gatherWithConcurrency(*userPermissionFutures)
        if any(userPermissions.error is not None for userPermissions in userPermissionsResults):
            await self.abortAndCleanup(f"Get user permissions failed for one or more controllers. Aborting.")

        anyUserNotAdmin = False
        for idx, userPermissions in enumerate(userPermissionsResults):
            adminRole = next(
                (role for role in userPermissions.data["roles"] if role["name"] == "super-admin"),
                None,
            )
            if adminRole is None:
                anyUserNotAdmin = True
                logging.error(f"{self.controllers[idx].host} - Login user does not have Account Owner role. Please modify permissions.")

        if anyUserNotAdmin:
            await self.abortAndCleanup(f"Login user not admin on one or more controllers. Aborting.")

        logging.info(f"----------Initializing Controller Data----------")
        # Gather applications across all controllers
        logging.info("Gathering Applications")
        getApplicationsFutures = [controller.getApplications() for controller in self.controllers]
        apmApplications = await gatherWithConcurrency(*getApplicationsFutures)
        if any(application.error is not None for application in apmApplications):
            await self.abortAndCleanup(f"One or more APM applications returned an error when fetching application list. Aborting.")
        apmApplications = [application.data for application in apmApplications]

        # Gather all controller configurations
        logging.info("Gathering Controller Configurations")
        getControllerConfigurations = [controller.getConfigurations() for controller in self.controllers]
        controllerConfigurations = await gatherWithConcurrency(*getControllerConfigurations)
        if any(controllerConfiguration.error is not None for controllerConfiguration in controllerConfigurations):
            await self.abortAndCleanup(f"One or more controller configurations returned an error. Aborting.")
        controllerConfigurations = [controllerConfiguration.data for controllerConfiguration in controllerConfigurations]

        # Gather analytics enabled status
        logging.info("Gathering Analytics Enabled Status")
        getAnalyticsEnabledStatus = [controller.getAnalyticsEnabledStatusForAllApplications() for controller in self.controllers]
        analyticsEnabledStatus = await gatherWithConcurrency(*getAnalyticsEnabledStatus)
        if any(analyticsEnabledStatus.error is not None for analyticsEnabledStatus in analyticsEnabledStatus):
            await self.abortAndCleanup(f"One or more analytics enabled status returned an error. Aborting.")
        analyticsEnabledStatus = [analyticsEnabledStatus.data for analyticsEnabledStatus in analyticsEnabledStatus]

        # Gather license usage
        logging.info("Gathering License Usage")
        getLicenseUsage = [controller.getAccountUsageSummary() for controller in self.controllers]
        licenseUsage = await gatherWithConcurrency(*getLicenseUsage)
        if any(licenseUsage.error is not None for licenseUsage in licenseUsage):
            await self.abortAndCleanup(f"One or more license usages returned an error. Aborting.")
        licenseUsage = [licenseUsage.data for licenseUsage in licenseUsage]

        # Gather all dashboards
        logging.info("Gathering Dashboards")
        getDashboards = [controller.getDashboards() for controller in self.controllers]
        dashboards = await asyncio.gather(*getDashboards)
        if any(dashboard.error is not None for dashboard in dashboards):
            await self.abortAndCleanup(f"One or more dashboard returned an error. Aborting.")
        dashboards = [dashboard.data for dashboard in dashboards]

        # Gather App Server Agent List
        logging.info("Gathering App Server Agent Agent List")
        getAppServerAgents = [controller.getAppServerAgents() for controller in self.controllers]
        appServerAgents = await gatherWithConcurrency(*getAppServerAgents)
        if any(appServerAgent.error is not None for appServerAgent in appServerAgents):
            await self.abortAndCleanup(f"One or more App Server Agents returned an error. Aborting.")
        appServerAgents = [appServerAgent.data for appServerAgent in appServerAgents]

        # Gather Machine Agent List
        logging.info("Gathering Machine Agent Agent List")
        getMachineAgents = [controller.getMachineAgents() for controller in self.controllers]
        machineAgents = await gatherWithConcurrency(*getMachineAgents)
        if any(machineAgent.error is not None for machineAgent in machineAgents):
            await self.abortAndCleanup(f"One or more Machine Agents returned an error. Aborting.")
        machineAgents = [machineAgent.data for machineAgent in machineAgents]

        # Gather Database Agent List
        logging.info("Gathering Database Agent Agent List")
        getDBAgents = [controller.getDBAgents() for controller in self.controllers]
        dbAgents = await gatherWithConcurrency(*getDBAgents)
        if any(dbAgent.error is not None for dbAgent in dbAgents):
            await self.abortAndCleanup(f"One or more DB Agents returned an error. Aborting.")
        dbAgents = [dbAgent.data for dbAgent in dbAgents]

        # Gather Analytics Agent List
        logging.info("Gathering Analytics Agent List")
        getAnalyticsAgents = [controller.getAnalyticsAgents() for controller in self.controllers]
        analyticsAgents = await gatherWithConcurrency(*getAnalyticsAgents)
        if any(analyticsAgent.error is not None for analyticsAgent in analyticsAgents):
            await self.abortAndCleanup(f"One or more Analytics Agents returned an error. Aborting.")
        analyticsAgents = [analyticsAgent.data for analyticsAgent in analyticsAgents]

        # Construct application information dictionary
        logging.info("Constructing Controller Data Dictionary")
        for idx, controller in enumerate(self.controllers):
            self.controllerData[controller.host] = OrderedDict()
            hostData = self.controllerData[controller.host]
            hostData["controller"] = controller
            hostData["configurations"] = controllerConfigurations[idx]
            hostData["analyticsEnabledStatus"] = analyticsEnabledStatus[idx]
            hostData["exportedDashboards"] = dashboards[idx]
            hostData["licenseUsage"] = licenseUsage[idx]

            hostData["appServerAgents"] = appServerAgents[idx]
            hostData["machineAgents"] = machineAgents[idx]
            hostData["dbAgents"] = dbAgents[idx]
            hostData["analyticsAgents"] = analyticsAgents[idx]

            hostData["apm"] = OrderedDict()
            hostData["dashboards"] = OrderedDict()
            hostData["containers"] = OrderedDict()
            hostData["brum"] = OrderedDict()
            hostData["mrum"] = OrderedDict()
            hostData["analytics"] = OrderedDict()

            for apmApplication in apmApplications[idx]:
                self.controllerData[controller.host]["apm"][apmApplication["name"]] = apmApplication

    def validateThresholdsFile(self):
        logging.info(f"----------Input Validation----------")
        logging.info(f"Validating Thresholds - {self.thresholdsFileName}")

        def thresholdStrictlyDecreasing(jobStep, thresholdMetric) -> bool:
            thresholds = self.thresholds[jobStep]
            if all(
                value is True
                for value in [
                    thresholds["platinum"][thresholdMetric],
                    thresholds["gold"][thresholdMetric],
                    thresholds["silver"][thresholdMetric],
                ]
            ):
                return True
            if all(
                value is False
                for value in [
                    thresholds["platinum"][thresholdMetric],
                    thresholds["gold"][thresholdMetric],
                    thresholds["silver"][thresholdMetric],
                ]
            ):
                return False
            return thresholds["platinum"][thresholdMetric] >= thresholds["gold"][thresholdMetric] >= thresholds["silver"][thresholdMetric]

        def thresholdStrictlyIncreasing(jobStep, thresholdMetric) -> bool:
            thresholds = self.thresholds[jobStep]
            if all(
                value is False
                for value in [
                    thresholds["platinum"][thresholdMetric],
                    thresholds["gold"][thresholdMetric],
                    thresholds["silver"][thresholdMetric],
                ]
            ):
                return True
            if all(
                value is True
                for value in [
                    thresholds["platinum"][thresholdMetric],
                    thresholds["gold"][thresholdMetric],
                    thresholds["silver"][thresholdMetric],
                ]
            ):
                return False
            return thresholds["platinum"][thresholdMetric] <= thresholds["gold"][thresholdMetric] <= thresholds["silver"][thresholdMetric]

        thresholdLevels = ["platinum", "gold", "silver"]

        fail = False
        for jobStep, currThresholdLevels in self.thresholds.items():
            if list(currThresholdLevels.keys()) != thresholdLevels:
                logging.error(f"Thresholds file does not contains all of {thresholdLevels} on JobStep {jobStep}")
                fail = True
                break
            if not (currThresholdLevels["platinum"].keys() == currThresholdLevels["gold"].keys() == currThresholdLevels["silver"].keys()):
                logging.error(f"Thresholds file does not contain same evaluation metrics for each threshold level on JobStep {jobStep}")
                fail = True
                break
            self.thresholds[jobStep]["direction"] = {}  # increasing or decreasing
            for metric in currThresholdLevels["platinum"].keys():
                if thresholdStrictlyIncreasing(jobStep, metric):
                    self.thresholds[jobStep]["direction"][metric] = "increasing"
                elif thresholdStrictlyDecreasing(jobStep, metric):
                    self.thresholds[jobStep]["direction"][metric] = "decreasing"
                else:
                    logging.error(
                        f"Thresholds file does not contain strictly increasing/decreasing evaluation metric thresholds for {metric} on JobStep {jobStep}"
                    )
                    fail = True
        if fail:
            logging.error(f"Invalid thresholds file. Aborting.")
            sys.exit(0)
        else:
            logging.debug(f"Validated thresholds file")

    async def process(self):
        logging.info(f"----------Extract----------")
        for jobStep in [*self.bsgSteps, *self.otherSteps]:
            await jobStep.extract(self.controllerData)

        logging.info(f"----------Analyze----------")
        for jobStep in [*self.bsgSteps, *self.otherSteps]:
            jobStep.analyze(self.controllerData, self.thresholds)

        logging.info(f"----------Report----------")
        for report in self.reports:
            report.createWorkbook(self.bsgSteps, self.controllerData, self.jobFileName)

    def finalize(self):
        now = int(time.time())
        with open(f"output/{self.jobFileName}/info.json", "w", encoding="utf-8") as f:
            json.dump(
                {"lastRun": now, "thresholds": self.thresholdsFileName},
                fp=f,
                ensure_ascii=False,
                indent=4,
            )

        with open(f"output/{self.jobFileName}/controllerData.json", "w", encoding="utf-8") as f:
            json.dump(
                self.controllerData,
                fp=f,
                default=jsonEncoder,
                ensure_ascii=False,
                indent=4,
            )

    async def abortAndCleanup(self, msg: str, error=True):
        """Closes open controller connections"""
        await gatherWithConcurrency(*[controller.close() for controller in self.controllers])
        if error:
            logging.error(msg)
            sys.exit(1)
        else:
            logging.info(msg)
            sys.exit(0)
