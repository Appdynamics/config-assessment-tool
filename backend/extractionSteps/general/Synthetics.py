import logging
from collections import OrderedDict

from api.appd.AppDService import AppDService
from extractionSteps.JobStepBase import JobStepBase
from util.asyncio_utils import AsyncioUtils


class Synthetics(JobStepBase):
    def __init__(self):
        super().__init__("brum")

    async def extract(self, controllerData):
        """
        Extract Synthetics details.
        1. Makes one API call per application to get Synthetic jobs.
        """
        jobStepName = type(self).__name__

        for host, hostInfo in controllerData.items():
            logging.info(f'{hostInfo["controller"].host} - Extracting {jobStepName}')
            controller: AppDService = hostInfo["controller"]

            getSyntheticJobFutures = []
            for application in hostInfo[self.componentType].values():
                getSyntheticJobFutures.append(controller.getSyntheticJobs(application["id"]))
            syntheticJobs = await AsyncioUtils.gatherWithConcurrency(*getSyntheticJobFutures)

            allSyntheticBillableTimes = []
            allSyntheticPrivateAgentUtilization = []
            for idx, application in enumerate(hostInfo[self.componentType].values()):
                scheduleIds = [job["config"]["id"] for job in syntheticJobs[idx].data["jobListDatas"]]
                allSyntheticBillableTimes.append(controller.getSyntheticBillableTime(application["id"], scheduleIds))
                if len(syntheticJobs[idx].data["jobListDatas"]):
                    configs = [job["config"] for job in syntheticJobs[idx].data["jobListDatas"]]
                    allSyntheticPrivateAgentUtilization.append(controller.getSyntheticPrivateAgentUtilization(application["id"], configs))

            syntheticBillableTimes = await AsyncioUtils.gatherWithConcurrency(*allSyntheticBillableTimes)
            syntheticPrivateAgentUtilization = await AsyncioUtils.gatherWithConcurrency(*allSyntheticPrivateAgentUtilization)

            syntheticBillableTimesMap = {}
            for syntheticBillableTime in syntheticBillableTimes:
                for syntheticBillableTimeData in syntheticBillableTime.data:
                    syntheticBillableTimesMap[syntheticBillableTimeData["scheduleId"]] = syntheticBillableTimeData

            syntheticPrivateAgentUtilizationMap = {}
            for syntheticPrivateAgentUtilizationData in syntheticPrivateAgentUtilization:
                for syntheticPrivateAgentUtilizationData in syntheticPrivateAgentUtilizationData.data:
                    syntheticPrivateAgentUtilizationMap[syntheticPrivateAgentUtilizationData["id"]] = syntheticPrivateAgentUtilizationData

            for idx, applicationName in enumerate(hostInfo[self.componentType]):
                application = hostInfo[self.componentType][applicationName]

                application["syntheticJobs"] = syntheticJobs[idx].data
                for job in application["syntheticJobs"]["jobListDatas"]:
                    if job["hasPrivateAgent"]:
                        job["privateAgentUtilization"] = syntheticPrivateAgentUtilizationMap[job["config"]["id"]]
                        job["billableTime"] = None
                    else:
                        job["billableTime"] = syntheticBillableTimesMap[job["config"]["id"]]
                        job["privateAgentUtilization"] = None

    def analyze(self, controllerData, thresholds):
        pass
