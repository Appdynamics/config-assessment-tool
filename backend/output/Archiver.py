import logging
import os
import shutil
from datetime import datetime

import pandas as pd
import xlsxwriter

from output.PostProcessReport import PostProcessReport

light_font = '#FFFFFF'
dark_font = '#000000'

medium_bg = '#A0A0A0'
dark_bg = '#000000'


class Archiver(PostProcessReport):

    def __init__(self):
        self.workbook = None
        self.analysis_sheet = None

    async def post_process(self, jobFileName):
        logging.info(f"Archiving generated report for job: {jobFileName}")

        # Define source and archive directories
        source_directory = f"output/{jobFileName}"
        archive_directory = "output/archive"

        # Ensure the archive directory exists
        if not os.path.exists(archive_directory):
            os.makedirs(archive_directory)
            logging.info(f"Created archive directory: {archive_directory}")

        # Get current timestamp in the format mmddyyyy_HH_SS
        timestamp = datetime.now().strftime("%m%d%Y_%H_%S")

        # Iterate through files in the source directory
        if os.path.exists(source_directory):
            for file_name in os.listdir(source_directory):
                # Skip `controllerData.json` and any files starting with `info`
                if file_name == "controllerData.json" or file_name.startswith("info"):
                    logging.info(f"Skipping file: {file_name}")
                    continue

                source_file_path = os.path.join(source_directory, file_name)

                # Skip directories, only process files
                if os.path.isfile(source_file_path):
                    # Construct the new filename with the timestamp
                    base_name, ext = os.path.splitext(file_name)
                    new_file_name = f"{base_name}_{timestamp}{ext}"

                    # Destination file path
                    destination_file_path = os.path.join(archive_directory, new_file_name)

                    # Copy file to the archive directory
                    shutil.copy2(source_file_path, destination_file_path)
                    logging.info(f"Archiving a backup copy at: {destination_file_path}")
        else:
            logging.error(f"Source directory {source_directory} does not exist.")