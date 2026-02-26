import asyncio
import logging
import sys

import click
from backend.core.Engine import Engine
from backend.util.click_utils import coro
from backend.util.logging_utils import initLogging


@click.command()
@click.option("-j", "--job-file", default="DefaultJob")
@click.option("-t", "--thresholds-file", default="DefaultThresholds")
@click.option("-d", "--debug", is_flag=True)
@click.option("-c", "--concurrent-connections", type=int)
@click.option("-u", "--username", default=None, hidden=True)
@click.option("-p", "--password", default=None, hidden=True)
@click.option("-a", "--auth-method", default=None, hidden=True)
@coro
async def main(job_file: str, thresholds_file: str, debug, concurrent_connections: int, username: str, password: str, auth_method: str):
    initLogging(debug)
    engine = Engine(job_file, thresholds_file, concurrent_connections, username, password, auth_method)
    await engine.run()


if __name__ == "__main__":
    """
    Generate Health Score of AppDynamics applications
    """
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())
