from abc import ABC, abstractmethod


class PostProcessReport(ABC):
    @abstractmethod
    async def post_process(self, report_names):
        """
        execute post processing reports
        """
        pass
