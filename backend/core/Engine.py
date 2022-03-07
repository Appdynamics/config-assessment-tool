import json
import logging
import math
import os
import sys
import time
import traceback
from collections import OrderedDict
from pathlib import Path

from api.appd.AppDService import AppDService
from extractionSteps.maturityAssessment.apm.ApmDashboards import ApmDashboards
from extractionSteps.maturityAssessment.apm.AppAgents import AppAgents
from extractionSteps.maturityAssessment.apm.Backends import Backends
from extractionSteps.maturityAssessment.apm.BusinessTransactions import BusinessTransactions
from extractionSteps.maturityAssessment.apm.DataCollectors import DataCollectors
from extractionSteps.maturityAssessment.apm.ErrorConfiguration import ErrorConfiguration
from extractionSteps.maturityAssessment.apm.HealthRulesAndAlertingAPM import HealthRulesAndAlertingAPM
from extractionSteps.maturityAssessment.apm.MachineAgents import MachineAgents
from extractionSteps.maturityAssessment.apm.OverallAssessmentAPM import OverallAssessmentAPM
from extractionSteps.maturityAssessment.apm.Overhead import Overhead
from extractionSteps.maturityAssessment.apm.ServiceEndpoints import ServiceEndpoints
from extractionSteps.general.ControllerLevelDetails import ControllerLevelDetails
from extractionSteps.general.CustomMetrics import CustomMetrics
from extractionSteps.maturityAssessment.brum.HealthRulesAndAlertingBRUM import HealthRulesAndAlertingBRUM
from extractionSteps.maturityAssessment.brum.NetworkRequests import NetworkRequests
from extractionSteps.maturityAssessment.brum.OverallAssessmentBrum import OverallAssessmentBRUM
from reports.reports.AgentMatrixReport import AgentMatrixReport
from reports.reports.MaturityAssessmentReport import MaturityAssessmentReport
from reports.reports.CustomMetricsReport import CustomMetricsReport
from reports.reports.LicenseReport import LicenseReport
from reports.reports.MaturityAssessmentReportRaw import RawMaturityAssessmentReport
from util.asyncio_utils import gatherWithConcurrency
from util.stdlib_utils import jsonEncoder


class Engine:
    def __init__(self, jobFileName: str, thresholdsFileName: str):
        logging.info(f'\n{open(f"backend/resources/img/splash.txt").read()}')
        self.codebaseVersion = open(f"VERSION").read()
        logging.info(f"Software Version: {self.codebaseVersion}")

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
                useProxy=controller.get("useProxy", False),
            )
            for controller in self.job
        ]
        self.controllerData = OrderedDict()
        self.otherSteps = [
            ControllerLevelDetails(),
            CustomMetrics(),
        ]
        self.maturityAssessmentSteps = [
            # APM Report
            AppAgents(),
            MachineAgents(),
            BusinessTransactions(),
            Backends(),
            Overhead(),
            ServiceEndpoints(),
            ErrorConfiguration(),
            HealthRulesAndAlertingAPM(),
            DataCollectors(),
            ApmDashboards(),
            OverallAssessmentAPM(),
            # BRUM Report
            NetworkRequests(),
            HealthRulesAndAlertingBRUM(),
            OverallAssessmentBRUM(),
        ]
        self.reports = [
            MaturityAssessmentReport(),
            RawMaturityAssessmentReport(),
            AgentMatrixReport(),
            CustomMetricsReport(),
            LicenseReport(),
        ]

    async def run(self):
        startTime = time.monotonic()

        try:
            await self.validateThresholdsFile()
            await self.initControllers()
            await self.process()
            self.finalize()
        except Exception as e:
            # catch exceptions here, so we can terminate coroutines before program exit
            logging.error("".join(traceback.TracebackException.from_exception(e).format()))

        logging.info(f"----------Complete----------")

        sizeBytes = os.path.getsize(f"output/{self.jobFileName}/controllerData.json")
        sizeName = ("B", "KB", "MB", "GB")
        i = int(math.floor(math.log(sizeBytes, 1024)))
        p = math.pow(1024, i)
        size = round(sizeBytes / p, 2)

        executionTime = time.monotonic() - startTime
        mins, secs = divmod(executionTime, 60)
        hours, mins = divmod(mins, 60)
        if hours > 0:
            executionTimeString = f"{int(hours)}h {int(mins)}m {int(secs)}s"
        elif mins > 0:
            executionTimeString = f"{int(mins)}m {int(secs)}s"
        else:
            executionTimeString = f"{int(secs)}s"

        totalCalls = sum([controller.totalCallsProcessed for controller in self.controllers])

        logging.info(f"Total API calls made: {totalCalls}")
        logging.info(f"Size of data retrieved: {size} {sizeName[i]}")
        logging.info(f"Total execution time: {executionTimeString}")

        await self.abortAndCleanup(
            "Exiting.",
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

        for idx, controller in enumerate(self.controllers):
            self.controllerData[controller.host] = OrderedDict()
            hostData = self.controllerData[controller.host]
            hostData["controller"] = controller

    async def validateThresholdsFile(self):
        logging.info(f"----------Input Validation----------")
        logging.info(f"Validating Thresholds - {self.thresholdsFileName}")

        if "version" not in self.thresholds:
            await self.abortAndCleanup(
                f"Thresholds file is not versioned. Please use thresholds file compatible with {self.codebaseVersion}. Aborting."
            )
        if self.codebaseVersion != self.thresholds["version"]:
            await self.abortAndCleanup(
                f"Thresholds file version {self.thresholds['version']} is incompatible with codebase version {self.codebaseVersion}. Aborting."
            )
        # only need this once right here, we remove it for simpler iteration of thresholds
        del self.thresholds["version"]

        def thresholdStrictlyDecreasing(jobStep, thresholdMetric, componentType: str) -> bool:
            thresholds = self.thresholds[componentType][jobStep]
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

        def thresholdStrictlyIncreasing(jobStep, thresholdMetric, componentType: str) -> bool:
            thresholds = self.thresholds[componentType][jobStep]
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
        for componentType, thresholds in self.thresholds.items():
            for jobStep, currThresholdLevels in thresholds.items():
                if list(currThresholdLevels.keys()) != thresholdLevels:
                    logging.error(f"Thresholds file does not contains all of {thresholdLevels} on JobStep {jobStep}")
                    fail = True
                    break
                if not (currThresholdLevels["platinum"].keys() == currThresholdLevels["gold"].keys() == currThresholdLevels["silver"].keys()):
                    logging.error(f"Thresholds file does not contain same evaluation metrics for each threshold level on JobStep {jobStep}")
                    fail = True
                    break
                thresholds[jobStep]["direction"] = {}  # increasing or decreasing
                for metric in currThresholdLevels["platinum"].keys():
                    if thresholdStrictlyIncreasing(jobStep, metric, componentType):
                        thresholds[jobStep]["direction"][metric] = "increasing"
                    elif thresholdStrictlyDecreasing(jobStep, metric, componentType):
                        thresholds[jobStep]["direction"][metric] = "decreasing"
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
        for jobStep in [*self.otherSteps, *self.maturityAssessmentSteps]:
            await jobStep.extract(self.controllerData)

        logging.info(f"----------Analyze----------")
        for jobStep in [*self.maturityAssessmentSteps, *self.otherSteps]:
            jobStep.analyze(self.controllerData, self.thresholds)

        logging.info(f"----------Report----------")
        for report in self.reports:
            report.createWorkbook(self.maturityAssessmentSteps, self.controllerData, self.jobFileName)

    def finalize(self):
        now = int(time.time())
        with open(f"output/{self.jobFileName}/info.json", "w", encoding="ISO-8859-1") as f:
            json.dump(
                {"lastRun": now, "thresholds": self.thresholdsFileName},
                fp=f,
                ensure_ascii=False,
                indent=4,
            )

        with open(f"output/{self.jobFileName}/controllerData.json", "w", encoding="ISO-8859-1") as f:
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
