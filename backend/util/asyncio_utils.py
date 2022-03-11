import asyncio
import logging


class AsyncioUtils:
    concurrentConnections = 50

    @staticmethod
    def init(concurrentConnections: int = 50):
        if concurrentConnections > 100:
            logging.warning(f"Concurrent connections ({concurrentConnections}) is too high. Setting to 100.")
            concurrentConnections = 100
        elif concurrentConnections < 1:
            logging.warning(f"Concurrent connections ({concurrentConnections}) is too low. Setting to 1.")
            concurrentConnections = 1
        else:
            logging.info(f"Setting concurrent connections to {concurrentConnections}.")
        AsyncioUtils.concurrentConnections = concurrentConnections

    @staticmethod
    async def gatherWithConcurrency(*tasks):
        semaphore = asyncio.Semaphore(AsyncioUtils.concurrentConnections)

        async def semTask(task):
            async with semaphore:
                return await task

        return await asyncio.gather(*[semTask(task) for task in tasks])
