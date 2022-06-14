import logging
from collections import OrderedDict

from api.appd.AppDService import AppDService
from extractionSteps.JobStepBase import JobStepBase


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
            for brumApplication in (await controller.getEumApplications()).data:
                hostInfo["brum"][brumApplication["name"]] = brumApplication
            logging.info(f'{hostInfo["controller"].host} - MRUM Applications')
            for mrumApplicationGroup in (await controller.getMRUMApplications()).data:
                for mrumApplication in mrumApplicationGroup["children"]:
                    mrumApplication["name"] = mrumApplication["internalName"]
                    hostInfo["mrum"][f"{mrumApplicationGroup['appKey']}-{mrumApplication['name']}"] = mrumApplication

            logging.info(f'{hostInfo["controller"].host} - Extracting Controller Configurations')
            hostInfo["configurations"] = (await controller.getConfigurations()).data
            hostInfo["analyticsEnabledStatus"] = (await controller.getAnalyticsEnabledStatusForAllApplications()).data

            logging.info(f'{hostInfo["controller"].host} - Extracting Dashboards')
            hostInfo["exportedDashboards"] = (await controller.getDashboards()).data

            logging.info(f'{hostInfo["controller"].host} - Extracting Licenses')
            hostInfo["accountLicenseUsage"] = (await controller.getAccountUsageSummary()).data
            hostInfo["eumLicenseUsage"] = (await controller.getEumLicenseUsage()).data

            logging.info(f'{hostInfo["controller"].host} - Extracting Agent Details')
            hostInfo["appServerAgents"] = (await controller.getAppServerAgents()).data
            hostInfo["machineAgents"] = (await controller.getMachineAgents()).data
            hostInfo["dbAgents"] = (await controller.getDBAgents()).data
            hostInfo["analyticsAgents"] = (await controller.getAnalyticsAgents()).data

    def analyze(self, controllerData, thresholds):
        pass
