import asyncio


async def gatherWithConcurrency(*tasks, size: int = 50):
    semaphore = asyncio.Semaphore(size)

    async def semTask(task):
        async with semaphore:
            return await task

    return await asyncio.gather(*[semTask(task) for task in tasks])
