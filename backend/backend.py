import asyncio
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
@click.option("-u", "--username", default=None, help="Adds the option to put username dynamically")
@click.option("-p", "--password", default=None, help="Adds the option to put password dynamically")
@click.option("--car", is_flag=True, help="Run configration analysis report")
@coro
async def main(job_file: str, thresholds_file: str, debug, concurrent_connections: int, username: str, password: str, car: bool):
    initLogging(debug)
    engine = Engine(job_file, thresholds_file, concurrent_connections, username, password, car)
    await engine.run()


if __name__ == "__main__":
    """
    Generate Health Score of AppDynamics applications
    """
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())
