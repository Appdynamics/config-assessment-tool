import asyncio
import logging
import sys

import click

from core.Engine import Engine
from util.click_utils import coro
from util.logging_utils import initLogging


@click.command()
@click.option("-j", "--job-file", default="DefaultJob")
@click.option("-t", "--thresholds-file", default="DefaultThresholds")
@click.option("-d", "--debug", is_flag=True)
@click.option("-c", "--concurrent-connections", type=int)
@click.option("-p", "--password", default=None, help="Adds the option to put password dynamically")
@coro
async def main(job_file: str, thresholds_file: str, debug, concurrent_connections: int, password: str):
    initLogging(debug)
    engine = Engine(job_file, thresholds_file, concurrent_connections, password)
    await engine.run()
    logging.info(debug)
    logging.info(password)


if __name__ == "__main__":
    """
    Generate Health Score of AppDynamics applications
    """
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())
