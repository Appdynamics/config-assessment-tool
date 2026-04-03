"""
logging_config.py
-----------------
This module sets up the logging configuration for the application.

Purpose:
- Configures logging to output messages to both the console and a rotating log file.
- Ensures that logs are properly formatted and stored for debugging and monitoring purposes.

Key Features:
- Creates a `logs` directory if it doesn't already exist.
- Logs messages to a file (`application.log`) with rotation:
  - Maximum file size: 5 MB.
  - Keeps up to 5 backup log files.
- Logs messages to the console for real-time monitoring.
- Formats log messages with timestamps, log levels, and logger names.
- Sets the default logging level to `INFO`.

Usage:
- Call `setup_logging()` at the start of the application to initialize logging.
"""

# logging_config.py
import logging
from logging.handlers import RotatingFileHandler
import os

def setup_logging():
    # Create logs directory if it doesn't exist
    logs_dir = "logs"
    os.makedirs(logs_dir, exist_ok=True)

    # Define log file path
    log_file = os.path.join(logs_dir, "application.log")

    # Create handlers
    stream_handler = logging.StreamHandler()  # Logs to console
    file_handler = RotatingFileHandler(log_file, maxBytes=5_000_000, backupCount=5)  # Logs to file

    # Define log format
    log_format = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    # Apply the format to both handlers
    stream_handler.setFormatter(log_format)
    file_handler.setFormatter(log_format)

    # Get the root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)  # Set the logging level

    # Add handlers to the root logger
    root_logger.addHandler(stream_handler)
    root_logger.addHandler(file_handler)
