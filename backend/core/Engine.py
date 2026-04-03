import json
import logging
import math
import os
import sys
import time
import asyncio
import traceback
from collections import OrderedDict
from pathlib import Path
import importlib.util

import requests

from backend.api.appd.AppDService import AppDService
from backend.api.appd.AuthMethod import AuthMethod
from backend.extractionSteps.general.ControllerLevelDetails import ControllerLevelDetails
from backend.extractionSteps.general.CustomMetrics import CustomMetrics
from backend.extractionSteps.general.Synthetics import Synthetics
from backend.extractionSteps.maturityAssessment.apm.AppAgentsAPM import AppAgentsAPM
from backend.extractionSteps.maturityAssessment.apm.BackendsAPM import BackendsAPM
from backend.extractionSteps.maturityAssessment.apm.BusinessTransactionsAPM import BusinessTransactionsAPM
from backend.extractionSteps.maturityAssessment.apm.DashboardsAPM import DashboardsAPM
from backend.extractionSteps.maturityAssessment.apm.DataCollectorsAPM import DataCollectorsAPM
from backend.extractionSteps.maturityAssessment.apm.ErrorConfigurationAPM import ErrorConfigurationAPM
from backend.extractionSteps.maturityAssessment.apm.HealthRulesAndAlertingAPM import HealthRulesAndAlertingAPM
from backend.extractionSteps.maturityAssessment.apm.MachineAgentsAPM import MachineAgentsAPM
from backend.extractionSteps.maturityAssessment.apm.OverallAssessmentAPM import OverallAssessmentAPM
from backend.extractionSteps.maturityAssessment.apm.OverheadAPM import OverheadAPM
from backend.extractionSteps.maturityAssessment.apm.ServiceEndpointsAPM import ServiceEndpointsAPM
from backend.extractionSteps.maturityAssessment.brum.HealthRulesAndAlertingBRUM import HealthRulesAndAlertingBRUM
from backend.extractionSteps.maturityAssessment.brum.NetworkRequestsBRUM import NetworkRequestsBRUM
from backend.extractionSteps.maturityAssessment.brum.OverallAssessmentBRUM import OverallAssessmentBRUM
from backend.extractionSteps.maturityAssessment.mrum.HealthRulesAndAlertingMRUM import HealthRulesAndAlertingMRUM
from backend.extractionSteps.maturityAssessment.mrum.NetworkRequestsMRUM import NetworkRequestsMRUM
from backend.extractionSteps.maturityAssessment.mrum.OverallAssessmentMRUM import OverallAssessmentMRUM
from backend.output.Archiver import Archiver
from backend.output.PostProcessReport import PostProcessReport
# from output.presentations.cxPpt import createCxPpt
from backend.output.presentations.cxPptTemplate import createCxPpt as createCxPptTemplate
from backend.output.reports.AgentMatrixReport import AgentMatrixReport
from backend.output.reports.ConfigurationAnalysisReport import ConfigurationAnalysisReport
from backend.output.reports.CustomMetricsReport import CustomMetricsReport
from backend.output.reports.DashboardReport import DashboardReport
from backend.output.reports.LicenseReport import LicenseReport
from backend.output.reports.MaturityAssessmentReport import MaturityAssessmentReport
from backend.output.reports.MaturityAssessmentReportRaw import RawMaturityAssessmentReport
from backend.output.reports.SyntheticsReport import SyntheticsReport
from backend.util.asyncio_utils import AsyncioUtils
from backend.util.stdlib_utils import base64Decode, base64Encode, isBase64, jsonEncoder

logger = logging.getLogger(__name__.split('.')[-1])

class Engine:
    def __init__(self, jobFileName: str, thresholdsFileName: str, concurrentConnections: int, user_name: str, password: str, auth_method : str):

        # should we run the configuration analysis report in post-processing?
        self.controllers = []

        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            # running as executable bundle
            path = sys._MEIPASS
            logger.info(f"Running as executable bundle, using {path} as root directory")
            # When frozen, input/output are adjacent to the executable, outside _internal
            # But sys.executable usually points to the bundle main executable
            self.user_data_dir = os.path.dirname(sys.executable)
        else:
            # running from source/docker
            # cd to config-assessment-tool root directory
            path = os.path.realpath(f"{__file__}/../../..")
            logger.info(f"Running from source/docker, using {path} as root directory")
            self.user_data_dir = path

        os.chdir(path)

        self.input_dir = os.path.join(self.user_data_dir, "input")
        self.output_dir = os.path.join(self.user_data_dir, "output")

        logger.info(f'\n{open(f"backend/resources/img/splash.txt").read()}')
        self.codebaseVersion = open(f"VERSION").read().strip()
        logger.info(f"Running Software Version: {self.codebaseVersion}")


        # Validate jobFileName and thresholdFileName
        self.jobFileName = jobFileName
        self.thresholdsFileName = thresholdsFileName

        job_file_path = os.path.join(self.input_dir, "jobs", f"{self.jobFileName}.json")
        thresholds_file_path = os.path.join(self.input_dir, "thresholds", f"{self.thresholdsFileName}.json")
        job_output_dir = os.path.join(self.output_dir, self.jobFileName)

        if not Path(job_file_path).exists():
            logger.error(f"Job file {job_file_path} does not exit. Aborting.")
            sys.exit(1)
        if not Path(thresholds_file_path).exists():
            logger.error(f"Job file {thresholds_file_path} does not exit. Aborting.")
            sys.exit(1)
        if not os.path.exists(job_output_dir):
            os.makedirs(job_output_dir)
        self.job = json.loads(open(job_file_path).read())
        self.thresholds = json.loads(open(thresholds_file_path).read())

        try:
            response = requests.request(
                "GET",
                "https://api.github.com/repos/appdynamics/config-assessment-tool/tags",
                verify=all([job.get("verifySsl", True)] for job in self.job),
                timeout=5,
            )
            latestTag = None
            if not response.ok:
                logger.warning(f"Unable to get latest tag from https://api.github.com/repos/appdynamics/config-assessment-tool/tags")
            else:
                latestTag = json.loads(response.text)[0]["name"]
                if latestTag != self.codebaseVersion:
                    logger.warning(f"You are using an outdated version of the software. Current {self.codebaseVersion} Latest {latestTag}")
                    logger.warning("You can get the latest version from https://github.com/Appdynamics/config-assessment-tool/releases")
        except requests.exceptions.RequestException:
            logger.warning(f"Unable to get latest tag from https://api.github.com/repos/appdynamics/config-assessment-tool/tags")

        # Default concurrent connections to 10 for On-Premise controllers
        if any(job for job in self.job if "saas.appdynamics.com" not in job["host"]):
            logger.info(f"On-Premise controller detected. It is recommended to use a maximum of 10 concurrent connections.")
            concurrentConnections = 10 if concurrentConnections is None else concurrentConnections
        else:
            logger.info(f"SaaS controller detected. It is recommended to use a maximum of 50 concurrent connections.")
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
        with open(os.path.join(self.input_dir, "jobs", f"{jobFileName}.json"), "w", encoding="ISO-8859-1") as f:
            json.dump(
                self.job,
                fp=f,
                ensure_ascii=False,
                indent=4,
            )

        for controller in self.job:

            if controller.get("authType") is None:
                logger.warn(f'\'authType\' is not '
                             f'specified for host {controller["host"]} in '
                             f'the input job file. Will '
                             f'default to basic authentication for '
                             f'backward compatibility.')
                controller["authType"] = "basic"

            logger.debug(
                f'authenticationMethod: {controller["authType"]} '
                f'for host {controller["host"]}')

            authMethod = AuthMethod(
                # auth_method=controller["authType"],
                auth_method=auth_method if auth_method else controller["authType"],
                host=controller["host"],
                port=controller["port"],
                ssl=controller["ssl"],
                account=controller["account"],
                username=user_name if user_name else controller["username"],
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
            logger.info("Dynamic password change was used!")  # have confirmation, that it's working as intended
        else:
            logger.info("Using password from jobfile")

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
            await self.validateThresholdsFile()
            await self.initControllers()
            await self.process()
            await self.postProcess()
            await self.runPlugins()
            self.finalize(startTime)
        except Exception as e:
            # catch exceptions here, so we can terminate coroutines before program exit
            logger.error("".join(traceback.TracebackException.from_exception(e).format()))

        await self.abortAndCleanup(
            "JOB FINISHED. IF USING WEB UI, YOU MAY CLOSE THIS MODAL WINDOW.",
            error=False,
        )


    async def runPlugins(self):
        logger.info(f"----------Plugins----------")
        plugin_dir = os.path.join(self.user_data_dir, "plugins")
        if not os.path.exists(plugin_dir):
            logger.info(f"No plugins directory found at {plugin_dir}")
            return

        # List subdirectories in plugins dir (ignoring files)
        plugin_folders = [d for d in os.listdir(plugin_dir) if os.path.isdir(os.path.join(plugin_dir, d)) and not d.startswith('__')]

        if not plugin_folders:
            logger.info("No plugin directories found.")
            return

        context = {
            "jobFileName": self.jobFileName,
            "outputDir": os.path.join(self.output_dir, self.jobFileName),
            "controllerData": self.controllerData,
            "controllers": self.controllers
        }

        for plugin_name in plugin_folders:
            plugin_path = os.path.join(plugin_dir, plugin_name)
            main_file = os.path.join(plugin_path, "main.py")

            if not os.path.exists(main_file):
                logger.debug(f"Plugin directory '{plugin_name}' exists but missing 'main.py'. Skipping.")
                continue

            requirements_file = os.path.join(plugin_path, "requirements.txt")
            if os.path.exists(requirements_file):
                logger.info(f"Loading plugin: {plugin_name}. Standalone plugin detected. Will not run this plugin as its not part of CAT tool lifecycle. You may launch separately.")
                continue

            try:
                logger.info(f"Loading plugin: {plugin_name}")

                # Import the module dynamically from the path
                spec = importlib.util.spec_from_file_location(f"plugins.{plugin_name}.main", main_file)
                if spec is None or spec.loader is None:
                     logger.error(f"Could not load spec for plugin {plugin_name}")
                     continue

                module = importlib.util.module_from_spec(spec)
                sys.modules[f"plugins.{plugin_name}.main"] = module

                # Temporarily add plugin directory to sys.path to allow internal imports within the plugin
                sys.path.insert(0, plugin_path)
                try:
                    spec.loader.exec_module(module)
                except SystemExit:
                    logger.warning(f"Plugin {plugin_name} called sys.exit() during loading. Skipping.")
                    continue
                except Exception as e:
                    logger.error(f"Error loading plugin {plugin_name}: {e}")
                    continue
                finally:
                    # Clean up sys.path
                    if sys.path[0] == plugin_path:
                        sys.path.pop(0)

                # Add a logging filter to prefix plugin logs
                class PluginLogFilter(logging.Filter):
                    def __init__(self, plugin_name):
                        super().__init__()
                        self.prefix = f"({plugin_name}): "

                    def filter(self, record):
                        # Avoid double prefixing if the plugin name is already in the message (unlikely but safe)
                        if not str(record.msg).startswith(self.prefix):
                            record.msg = self.prefix + str(record.msg)
                        return True

                plugin_filter = PluginLogFilter(plugin_name)
                logging.getLogger().addFilter(plugin_filter)

                try:
                    logger.info(f"Executing plugin: {plugin_name}")
                    try:
                        # Check if the function is a coroutine (async)
                        if asyncio.iscoroutinefunction(module.run_plugin):
                            result = await module.run_plugin(context)
                        else:
                            result = module.run_plugin(context)

                        logger.info(f"Plugin {plugin_name} finished. Result: {result}")
                    except Exception as e:
                        logger.error(f"Error running plugin {plugin_name}: {e}")
                        logger.debug(traceback.format_exc())
                finally:
                    logging.getLogger().removeFilter(plugin_filter)

            except Exception as e:
                logger.error(f"Failed to load plugin {plugin_name}: {e}")
                logger.debug(traceback.format_exc())

    async def initControllers(self) -> ([AppDService], str):
        logger.info(f"Validating Controller Login(s) for Job - {self.jobFileName} ")
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
        logger.info(f"----------Input Validation----------")
        logger.info(f"Validating Thresholds - {self.thresholdsFileName}")

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
                    logger.error(f"Thresholds file does not contains all of {thresholdLevels} on JobStep {jobStep}")
                    fail = True
                    break
                if not (currThresholdLevels["platinum"].keys() == currThresholdLevels["gold"].keys() == currThresholdLevels["silver"].keys()):
                    logger.error(f"Thresholds file does not contain same evaluation metrics for each threshold level on JobStep {jobStep}")
                    fail = True
                    break
                thresholds[jobStep]["direction"] = {}  # increasing or decreasing
                for metric in currThresholdLevels["platinum"].keys():
                    if thresholdStrictlyIncreasing(jobStep, metric, componentType):
                        thresholds[jobStep]["direction"][metric] = "increasing"
                    elif thresholdStrictlyDecreasing(jobStep, metric, componentType):
                        thresholds[jobStep]["direction"][metric] = "decreasing"
                    else:
                        logger.error(
                            f"Thresholds file does not contain strictly increasing/decreasing evaluation metric thresholds for {metric} on JobStep {jobStep}"
                        )
                        fail = True
        if fail:
            logger.error(f"Invalid thresholds file. Aborting.")
            sys.exit(0)
        else:
            logger.debug(f"Validated thresholds file")

    async def process(self):
        logger.info(f"----------Extract----------")
        for jobStep in [*self.otherSteps, *self.maturityAssessmentSteps]:
            await jobStep.extract(self.controllerData)

        logger.info(f"----------Analyze----------")
        for jobStep in [*self.maturityAssessmentSteps, *self.otherSteps]:
            jobStep.analyze(self.controllerData, self.thresholds)

        logger.info(f"----------Report----------")
        job_output_dir = os.path.join(self.output_dir, self.jobFileName)
        for report in self.reports:
            report.createWorkbook(self.maturityAssessmentSteps, self.controllerData, self.jobFileName, self.output_dir)

    def finalize(self, startTime):
        now = int(time.time())
        job_output_dir = os.path.join(self.output_dir, self.jobFileName)

        with open(os.path.join(job_output_dir, "info.json"), "w", encoding="utf-8") as f:
            json.dump(
                {"lastRun": now, "thresholds": self.thresholdsFileName},
                fp=f,
                ensure_ascii=False,
                indent=4,
            )

        with open(os.path.join(job_output_dir, "controllerData.json"), "w", encoding="utf-8") as f:
            json.dump(
                self.controllerData,
                fp=f,
                default=jsonEncoder,
                indent=4,
            )

        # createCxPpt(self.jobFileName)
        createCxPptTemplate(self.jobFileName, self.output_dir)

        logger.info(f"----------Complete----------")
        # if controllerData.json file exists, delete it
        controller_data_path = os.path.join(job_output_dir, "controllerData.json")
        if Path(controller_data_path).exists():
            sizeBytes = os.path.getsize(controller_data_path)
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

            logger.info(f"Total API calls made: {totalCalls}")
            logger.info(f"Size of data retrieved: {size} {sizeName[i]}")
            logger.info(f"Total execution time: {executionTimeString}")

    async def abortAndCleanup(self, msg: str, error=True):
        """Closes open controller connections"""
        await AsyncioUtils.gatherWithConcurrency(*[controller.close() for controller in self.controllers])
        if error:
            logger.error(msg)
            sys.exit(1)
        else:
            if msg:
                logger.info(msg)
            sys.exit(0)

    async def postProcess(self):
        logger.info(f"----------Post Process----------")
        commands = []
        commands.append(ConfigurationAnalysisReport(self.output_dir))

        # after ALL reports generated archive a copy for safekeeping
        commands.append(Archiver(self.output_dir))

        for command in commands:
            await command.post_process(self.jobFileName)

        logger.info(f"----------Post Process Done----------")
