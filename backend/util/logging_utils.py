import asyncio
import logging
import os


def initLogging(debug: bool):
    """Set up logging."""
    # cd to config-assessment-tool root directory
    path = os.path.realpath(f"{__file__}/../../..")
    os.chdir(path)

    if not os.path.exists("logs"):
        os.makedirs("logs")

    logging.basicConfig(
        level=logging.DEBUG if debug else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler("logs/config-assessment-tool-backend.log"),
            logging.StreamHandler(),
        ],
    )
    EventLoopDelayMonitor()


class EventLoopDelayMonitor:
    def __init__(self, loop=None, start=True, interval=1, logger=None):
        self._interval = interval
        self._log = logger or logging.getLogger(__name__)
        self._loop = loop or asyncio.get_event_loop()
        if start:
            self.start()

    def run(self):
        self._loop.call_later(self._interval, self._handler, self._loop.time())

    def _handler(self, start_time):
        latency = (self._loop.time() - start_time) - self._interval

        self._log.debug(
            f"asyncio - Task count: {len(asyncio.all_tasks())} - EventLoop delay %.4f",
            latency,
        )

        for coro in asyncio.all_tasks():
            if not hasattr(coro.get_coro(), "cr_frame"):
                continue

            f_locals = coro.get_coro().cr_frame.f_locals

            obj = f_locals.get("self", None)
            host = obj.host if hasattr(obj, "host") else None

            debugString = f_locals.get("debugString", None)

            if host is not None and debugString is not None:
                self._log.debug(f"asyncio PENDING tasks - {host} - {debugString}")

        if not self.is_stopped():
            self.run()

    def is_stopped(self):
        return self._stopped

    def start(self):
        self._stopped = False
        self.run()

    def stop(self):
        self._stopped = True
