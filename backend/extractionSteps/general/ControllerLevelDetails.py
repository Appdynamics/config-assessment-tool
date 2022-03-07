import logging
from collections import OrderedDict

from api.appd.AppDService import AppDService
from extractionSteps.JobStepBase import JobStepBase
from util.asyncio_utils import gatherWithConcurrency


class ControllerLevelDetails(JobStepBase):
    def __init__(self):
        super().__init__("controller")

    async def extract(self, controllerData):
        """
        Extract controller level details.
        """
        jobStepName = type(self).__name__

        for host, hostInfo in controllerData.items():
            controller: AppDService = hostInfo["controller"]

            hostInfo["apm"] = OrderedDict()
            hostInfo["dashboards"] = OrderedDict()
            hostInfo["containers"] = OrderedDict()
            hostInfo["brum"] = OrderedDict()
            hostInfo["mrum"] = OrderedDict()
            hostInfo["analytics"] = OrderedDict()

            logging.info(f'{hostInfo["controller"].host} - Extracting APM Applications')
            for apmApplication in (await controller.getApmApplications()).data:
                hostInfo["apm"][apmApplication["name"]] = apmApplication
            logging.info(f'{hostInfo["controller"].host} - EUM Applications')
            for eumApplication in (await controller.getEumApplications()).data:
                hostInfo["brum"][eumApplication["name"]] = eumApplication

            logging.info(f'{hostInfo["controller"].host} - Extracting Controller Configurations')
            hostInfo["configurations"] = (await controller.getConfigurations()).data
            hostInfo["analyticsEnabledStatus"] = (await controller.getAnalyticsEnabledStatusForAllApplications()).data

            logging.info(f'{hostInfo["controller"].host} - Extracting Dashboards')
            hostInfo["exportedDashboards"] = (await controller.getDashboards()).data

            logging.info(f'{hostInfo["controller"].host} - Extracting Licenses')
            hostInfo["appServerAgents"] = (await controller.getAppServerAgents()).data
            hostInfo["machineAgents"] = (await controller.getMachineAgents()).data
            hostInfo["dbAgents"] = (await controller.getDBAgents()).data
            hostInfo["analyticsAgents"] = (await controller.getAnalyticsAgents()).data
            hostInfo["licenseUsage"] = (await controller.getAccountUsageSummary()).data

    def analyze(self, controllerData, thresholds):
        pass