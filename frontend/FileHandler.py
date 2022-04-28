import logging
import os
import subprocess
import sys

from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib import parse


def getPlatform():
    platform = sys.platform
    if sys.platform == "linux":
        proc_version = open("/proc/version").read()
        if "microsoft" in proc_version:
            platform = "wsl"
    return platform


def openFile(filename):
    logging.info("Opening file: " + filename)
    platform = getPlatform()

    if platform == "darwin":
        subprocess.call(("open", filename))
    elif platform in ["win64", "win32"]:
        os.startfile(filename.replace("/", "\\"))
    elif platform == "wsl":
        subprocess.call(["wslview", filename])
    else:  # linux variants
        subprocess.call(("xdg-open", filename))


def openFolder(path):
    logging.info("Opening folder: " + path)
    platform = getPlatform()

    if platform == "darwin":
        subprocess.call(["open", "--", path])
    elif platform in ["win64", "win32"]:
        subprocess.call(["start", path])
    elif platform == "wsl":
        command = "explorer.exe `wslpath -w " + path + "`"
        subprocess.run(["bash", "-c", command])
    else:  # linux variants
        subprocess.call(["xdg-open", "--", path])


class MyServer(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/html")
        self.end_headers()

        query_components = dict(parse.parse_qsl(parse.urlsplit(self.path).query))

        if self.path == "/ping":
            self.wfile.write(b"pong")
        elif "type" in query_components and query_components["type"] == "file":
            openFile(query_components["path"])
        elif "type" in query_components and query_components["type"] == "folder":
            openFolder(query_components["path"])

    def log_message(self, format, *args):
        logging.info("%s - - [%s] %s" % (self.address_string(), self.log_date_time_string(), format % args))


if __name__ == "__main__":
    if not os.path.exists("logs"):
        os.makedirs("logs")
    if not os.path.exists("output"):
        os.makedirs("output")

    # noinspection PyArgumentList
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler("logs/config-assessment-tool-frontend.log"),
        ],
    )

    hostName = "localhost"
    serverPort = 16225

    try:
        logging.info("Starting FileHandler on " + hostName + ":" + str(serverPort))
        webServer = HTTPServer((hostName, serverPort), MyServer)

        webServer.serve_forever()
    except KeyboardInterrupt:
        logging.info("Stopping FileHandler")
