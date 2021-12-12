import logging
from collections import OrderedDict

from api.appd.AppDService import AppDService
from extractionSteps.JobStepBase import BSGJobStepBase
from util.asyncio_utils import gatherWithConcurrency


class Overhead(BSGJobStepBase):
    def __init__(self):
        super().__init__("apm")

    async def extract(self, controllerData):
        """
        Extract node level details.
        1. Makes one API call per application to get Dev level Monitoring Configuration for Application (PRODUCTION or DEVELOPMENT).
        2. Makes one API call per application to get Dev level Monitoring Configuration per BT (enabled or disabled).
        3. Makes one API call per application to get Node Properties for application components which have been modified from the default (find-entry-points).
        3. Makes one API call per application to get application Call Graph Settings (aggressive snapshotting).
        """
        jobStepName = type(self).__name__

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Extracting details for {jobStepName}')
            controller: AppDService = hostInfo["controller"]

            # Gather necessary metrics.
            getDevModeConfigFutures = []
            getInstrumentationLevelFutures = []
            getAllNodePropertiesForCustomizedComponentsFutures = []
            getApplicationConfigurationFutures = []
            for application in hostInfo[self.componentType].values():
                getDevModeConfigFutures.append(controller.getDevModeConfig(application["id"]))
                getInstrumentationLevelFutures.append(controller.getInstrumentationLevel(application["id"]))
                getAllNodePropertiesForCustomizedComponentsFutures.append(controller.getAllNodePropertiesForCustomizedComponents(application["id"]))
                getApplicationConfigurationFutures.append(controller.getApplicationConfiguration(application["id"]))
            devModeConfigs = await gatherWithConcurrency(*getDevModeConfigFutures)
            instrumentationLevels = await gatherWithConcurrency(*getInstrumentationLevelFutures)
            nodePropertiesForCustomizedComponents = [
                component.data for component in await gatherWithConcurrency(*getAllNodePropertiesForCustomizedComponentsFutures)
            ]
            applicationConfigurationSettings = await gatherWithConcurrency(*getApplicationConfigurationFutures)

            for idx, applicationName in enumerate(hostInfo[self.componentType]):
                application = hostInfo[self.componentType][applicationName]
                application["agentConfigurations"] = [
                    component.data for component in nodePropertiesForCustomizedComponents[idx] if component.data != []
                ]
                application["devModeConfig"] = devModeConfigs[idx].data
                application["instrumentationLevel"] = instrumentationLevels[idx].data
                application["applicationConfiguration"] = applicationConfigurationSettings[idx].data

    def analyze(self, controllerData, thresholds):
        """
        Analysis of node level details.
        1. Determines if Developer Mode is either enabled application wide or for any BT.
        2. Determines if find-entry-points is enabled for any application component.
        3. Determines if Aggressive Snapshotting is enabled.
        """

        jobStepName = type(self).__name__

        # Get thresholds related to job
        jobStepThresholds = thresholds[jobStepName]

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Analyzing details for {jobStepName}')

            for application in hostInfo[self.componentType].values():
                # Root node of current application for current JobStep.
                analysisDataRoot = application[jobStepName] = OrderedDict()
                # This data goes into the 'JobStep - Metrics' xlsx sheet.
                analysisDataEvaluatedMetrics = analysisDataRoot["evaluated"] = OrderedDict()
                # This data goes into the 'JobStep - Raw' xlsx sheet.
                analysisDataRawMetrics = analysisDataRoot["raw"] = OrderedDict()

                # developerModeNotEnabledForAnyBT
                numberOfTransactionsWithDevModeEnabled = 0
                for config in application["devModeConfig"]:
                    if config["children"] is not None:
                        for child in config["children"]:
                            if child["enabled"]:
                                numberOfTransactionsWithDevModeEnabled += 1
                analysisDataEvaluatedMetrics["developerModeNotEnabledForAnyBT"] = numberOfTransactionsWithDevModeEnabled == 0

                # findEntryPointsNotEnabled
                numberOfComponentsWithFiendEntryPointsEnabled = 0
                for agentConfiguration in application["agentConfigurations"]:
                    if agentConfiguration is not None:
                        findEntryPointsProperty = next(
                            iter(
                                [
                                    property
                                    for property in agentConfiguration["properties"]
                                    if property
                                    if property["definition"]["name"] == "find-entry-points"
                                ]
                            ),
                            None,
                        )
                        if findEntryPointsProperty is not None and findEntryPointsProperty["stringValue"] == "true":
                            numberOfComponentsWithFiendEntryPointsEnabled += 1
                analysisDataEvaluatedMetrics["findEntryPointsNotEnabled"] = numberOfComponentsWithFiendEntryPointsEnabled == 0

                # aggressiveSnapshottingNotEnabled
                aggressiveSnapshottingNotEnabled = True
                for config, value in application["applicationConfiguration"].items():
                    if config.lower().endswith("callgraphconfiguration"):
                        if "hotspotsEnabled" in value and value["hotspotsEnabled"]:
                            aggressiveSnapshottingNotEnabled = False
                analysisDataEvaluatedMetrics["aggressiveSnapshottingNotEnabled"] = aggressiveSnapshottingNotEnabled

                # instrumentationLevel
                instrumentationLevel = application["instrumentationLevel"]
                analysisDataEvaluatedMetrics["developerModeNotEnabledForApplication"] = instrumentationLevel == "PRODUCTION"

                analysisDataRawMetrics["numberOfComponentsWithFiendEntryPointsEnabled"] = numberOfComponentsWithFiendEntryPointsEnabled
                analysisDataRawMetrics["numberOfTransactionsWithDevModeEnabled"] = numberOfTransactionsWithDevModeEnabled

                self.applyThresholds(analysisDataEvaluatedMetrics, analysisDataRoot, jobStepThresholds)
