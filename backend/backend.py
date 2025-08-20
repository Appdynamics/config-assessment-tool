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
@click.option("-u", "--username", default=None, help="overwrite job file with this username")
@click.option("-p", "--password", default=None, help="overwrite job file with this password")
@click.option("-m", "--auth-method", default=None, help="overwrite job file with this auth-method(basic,secret,token)")
@click.option("--car", is_flag=True, help="Generate the configration analysis report as part of the output")
@coro
async def main(job_file: str, thresholds_file: str, debug, concurrent_connections: int, username: str, password: str, auth_method: str,  car: bool):
    initLogging(debug)
    engine = Engine(job_file, thresholds_file, concurrent_connections, username, password, auth_method, car)
    await engine.run()


if __name__ == "__main__":
    """
    Generate Health Score of AppDynamics applications
    """
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())
