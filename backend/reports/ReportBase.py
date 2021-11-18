from abc import ABC, abstractmethod


class ReportBase(ABC):
    @abstractmethod
    async def createWorkbook(self, jobs, controllerData, jobFileName):
        """
        Extraction step of AppDynamics data.
        """
        pass
