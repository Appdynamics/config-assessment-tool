import json
import logging
import math
import os
import subprocess
import sys
import time
import traceback
import webbrowser
from collections import OrderedDict
from pathlib import Path

import requests
from flask import Flask
from numpy.typing.mypy_plugin import plugin

from api.appd.AppDService import AppDService
from api.appd.AuthMethod import AuthMethod
from extractionSteps.general.ControllerLevelDetails import ControllerLevelDetails
from extractionSteps.general.CustomMetrics import CustomMetrics
from extractionSteps.general.Synthetics import Synthetics
from extractionSteps.maturityAssessment.apm.AppAgentsAPM import AppAgentsAPM
from extractionSteps.maturityAssessment.apm.BackendsAPM import BackendsAPM
from extractionSteps.maturityAssessment.apm.BusinessTransactionsAPM import BusinessTransactionsAPM
from extractionSteps.maturityAssessment.apm.DashboardsAPM import DashboardsAPM
from extractionSteps.maturityAssessment.apm.DataCollectorsAPM import DataCollectorsAPM
from extractionSteps.maturityAssessment.apm.ErrorConfigurationAPM import ErrorConfigurationAPM
from extractionSteps.maturityAssessment.apm.HealthRulesAndAlertingAPM import HealthRulesAndAlertingAPM
from extractionSteps.maturityAssessment.apm.MachineAgentsAPM import MachineAgentsAPM
from extractionSteps.maturityAssessment.apm.OverallAssessmentAPM import OverallAssessmentAPM
from extractionSteps.maturityAssessment.apm.OverheadAPM import OverheadAPM
from extractionSteps.maturityAssessment.apm.ServiceEndpointsAPM import ServiceEndpointsAPM
from extractionSteps.maturityAssessment.brum.HealthRulesAndAlertingBRUM import HealthRulesAndAlertingBRUM
from extractionSteps.maturityAssessment.brum.NetworkRequestsBRUM import NetworkRequestsBRUM
from extractionSteps.maturityAssessment.brum.OverallAssessmentBRUM import OverallAssessmentBRUM
from extractionSteps.maturityAssessment.mrum.HealthRulesAndAlertingMRUM import HealthRulesAndAlertingMRUM
from extractionSteps.maturityAssessment.mrum.NetworkRequestsMRUM import NetworkRequestsMRUM
from extractionSteps.maturityAssessment.mrum.OverallAssessmentMRUM import OverallAssessmentMRUM
from output.PostProcessReport import PostProcessReport
from output.presentations.cxPpt import createCxPpt
from output.presentations.cxPptFsoUseCases import createCxHamUseCasePpt
from output.reports.AgentMatrixReport import AgentMatrixReport
from output.reports.ConfigurationAnalysisReport import ConfigurationAnalysisReport
from output.reports.CustomMetricsReport import CustomMetricsReport
from output.reports.DashboardReport import DashboardReport
from output.reports.LicenseReport import LicenseReport
from output.reports.MaturityAssessmentReport import MaturityAssessmentReport
from output.reports.MaturityAssessmentReportRaw import RawMaturityAssessmentReport
from output.reports.SyntheticsReport import SyntheticsReport
from util.asyncio_utils import AsyncioUtils
from util.stdlib_utils import base64Decode, base64Encode, isBase64, jsonEncoder


class Engine:
    def __init__(self, jobFileName: str, thresholdsFileName: str, concurrentConnections: int, username: str, password: str, car: bool, compare: bool):

        # should we run the configuration analysis report in post-processing?
        self.controllers = []
        self.car = car
        self.compare = compare

        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            # running as executable bundle
            path = sys._MEIPASS
            logging.info(f"Running as executable bundle, using {path} as root directory")
        else:
            # running from source/docker
            # cd to config-assessment-tool root directory
            path = os.path.realpath(f"{__file__}/../../..")
            logging.info(f"Running from source/docker, using {path} as root directory")

        os.chdir(path)

        logging.info(f'\n{open(f"backend/resources/img/splash.txt").read()}')
        self.codebaseVersion = open(f"VERSION").read()
        logging.info(f"Running Software Version: {self.codebaseVersion}")


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

        try:
            response = requests.request(
                "GET",
                "https://api.github.com/repos/appdynamics/config-assessment-tool/tags",
                verify=all([job.get("verifySsl", True)] for job in self.job),
                timeout=5,
            )
            latestTag = None
            if not response.ok:
                logging.warning(f"Unable to get latest tag from https://api.github.com/repos/appdynamics/config-assessment-tool/tags")
            else:
                latestTag = json.loads(response.text)[0]["name"]
                if latestTag != self.codebaseVersion:
                    logging.warning(f"You are using an outdated version of the software. Current {self.codebaseVersion} Latest {latestTag}")
                    logging.warning("You can get the latest version from https://github.com/Appdynamics/config-assessment-tool/releases")
        except requests.exceptions.RequestException:
            logging.warning(f"Unable to get latest tag from https://api.github.com/repos/appdynamics/config-assessment-tool/tags")

        # Default concurrent connections to 10 for On-Premise controllers
        if any(job for job in self.job if "saas.appdynamics.com" not in job["host"]):
            logging.info(f"On-Premise controller detected. It is recommended to use a maximum of 10 concurrent connections.")
            concurrentConnections = 10 if concurrentConnections is None else concurrentConnections
        else:
            logging.info(f"SaaS controller detected. It is recommended to use a maximum of 50 concurrent connections.")
            concurrentConnections = 50 if concurrentConnections is None else concurrentConnections
        AsyncioUtils.init(concurrentConnections)

        # Convert passwords to base64 if they aren't already
        for controller in self.job:
            if not isBase64(controller["pwd"]):
                # not base64, so encode it
                controller["pwd"] = base64Encode(f"CAT-ENCODED-{controller['pwd']}")
            elif not base64Decode(controller["pwd"]).startswith("CAT-ENCODED-"):
                # is valid base64, but doesn't contain our encoding string "CAT-ENCODED-"
                controller["pwd"] = base64Encode(f"CAT-ENCODED-{controller['pwd']}")

            # add in fields not present
            if "verifySsl" not in controller:
                controller["verifySsl"] = True
            if "useProxy" not in controller:
                controller["useProxy"] = True
            if "applicationFilter" not in controller:
                controller["applicationFilter"] = {"apm": ".*", "mrum": ".*", "brum": ".*"}
            if "timeRangeMins" not in controller:
                controller["timeRangeMins"] = 1440

        # Save the job back to disk with the updated password
        with open(f"input/jobs/{jobFileName}.json", "w", encoding="ISO-8859-1") as f:
            json.dump(
                self.job,
                fp=f,
                ensure_ascii=False,
                indent=4,
            )

        for controller in self.job:

            if controller.get("authType") is None:
                logging.warn(f'\'authType\' is not '
                             f'specified for host {controller["host"]} in '
                             f'the input job file. Will '
                             f'default to basic authentication for '
                             f'backward compatibility.')
                controller["authType"] = "basic"

            logging.debug(
                f'authenticationMethod: {controller["authType"]} '
                f'for host {controller["host"]}')

            authMethod = AuthMethod(
                auth_method=controller["authType"],
                host=controller["host"],
                port=controller["port"],
                ssl=controller["ssl"],
                account=controller["account"],
                username=username if username else controller["username"],
                password=password if password else base64Decode(controller[
                                                                    "pwd"])[len("CAT-ENCODED-") :],
                verifySsl=controller.get("verifySsl", True),
                useProxy=controller.get("useProxy", False)
            )

            controllerService = AppDService(
                applicationFilter=controller.get("applicationFilter", None),
                timeRangeMins=controller.get("timeRangeMins", 1440),
                authMethod=authMethod
            )


            self.controllers.append(controllerService)
            username = None
            password = None


        if password:  # I will let it here until it's the final version, so that we will
            logging.info("Dynamic password change was used!")  # have confirmation, that it's working as intended
        else:
            logging.info("Using password from jobfile")

        self.controllerData = OrderedDict()
        self.otherSteps = [
            ControllerLevelDetails(),
            CustomMetrics(),
            Synthetics(),
        ]
        self.maturityAssessmentSteps = [
            # APM Report
            AppAgentsAPM(),
            MachineAgentsAPM(),
            BusinessTransactionsAPM(),
            BackendsAPM(),
            OverheadAPM(),
            ServiceEndpointsAPM(),
            ErrorConfigurationAPM(),
            HealthRulesAndAlertingAPM(),
            DataCollectorsAPM(),
            DashboardsAPM(),
            OverallAssessmentAPM(),
            # BRUM Report
            NetworkRequestsBRUM(),
            HealthRulesAndAlertingBRUM(),
            OverallAssessmentBRUM(),
            # MRUM Report
            NetworkRequestsMRUM(),
            HealthRulesAndAlertingMRUM(),
            OverallAssessmentMRUM(),
        ]
        self.reports = [
            MaturityAssessmentReport(),
            RawMaturityAssessmentReport(),
            AgentMatrixReport(),
            CustomMetricsReport(),
            LicenseReport(),
            SyntheticsReport(),
            DashboardReport(),
        ]

    async def run(self):
        startTime = time.monotonic()

        try:
            self.preprocess()
            await self.validateThresholdsFile()
            await self.initControllers()
            await self.process()
            await self.postProcess()
            self.finalize(startTime)
        except Exception as e:
            # catch exceptions here, so we can terminate coroutines before program exit
            logging.error("".join(traceback.TracebackException.from_exception(e).format()))

        await self.abortAndCleanup(
            "",
            error=False,
        )

    async def initControllers(self) -> ([AppDService], str):
        logging.info(f"Validating Controller Login(s) for Job - {self.jobFileName} ")
        loginFutures = [controller.getAuthMethod().authenticate() for controller in
                        self.controllers]
        loginResults = await AsyncioUtils.gatherWithConcurrency(*loginFutures)
        if any(login.error is not None for login in loginResults):
            await self.abortAndCleanup(f"Unable to connect to one or more controllers. Aborting.")


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

    def finalize(self, startTime):
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

        createCxPpt(self.jobFileName)
        createCxHamUseCasePpt(self.jobFileName)

        logging.info(f"----------Complete----------")
        # if controllerData.json file exists, delete it
        if Path(f"output/{self.jobFileName}/controllerData.json").exists():
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

    async def abortAndCleanup(self, msg: str, error=True):
        """Closes open controller connections"""
        await AsyncioUtils.gatherWithConcurrency(*[controller.close() for controller in self.controllers])
        if error:
            logging.error(msg)
            sys.exit(1)
        else:
            if msg:
                logging.info(msg)
            sys.exit(0)

    async def postProcess(self):
        """Post-processing reports that rely on the core generated excel reports.
        Currently, for CAR(configuration analysis report) that consumes the core
        reports post process
         """
        logging.info(f"----------Post Process----------")
        commands = []

        if self.car:
            commands.append(ConfigurationAnalysisReport())

        for command in commands:
            if isinstance(command, PostProcessReport):
                await command.post_process(self.jobFileName)

        logging.info(f"----------Post Process Done----------")

    def preprocess(self):
        # launch compare app
        if self.compare:
            self.launch_flask_app("compare-plugin/core.py")
        logging.info(f"----------Preprocess Done----------")

    def launch_flask_app(self, script_path):
        script_abs_path = os.path.abspath(script_path)
        script_dir = os.path.dirname(script_abs_path)
        os.chdir(script_dir)
        logging.info(f"Changed working directory to: {script_dir}")
        app = Flask(__name__)
        url = "http://127.0.0.1:5000"  # This is the default URL where Flask serves
        webbrowser.open(url)
        app.run(debug=False)