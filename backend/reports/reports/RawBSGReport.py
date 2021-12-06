import logging

from openpyxl import Workbook

from reports.ReportBase import ReportBase


class RawDataReport(ReportBase):
    def createWorkbook(self, jobs, controllerData, jobFileName):
        logging.info(f"Creating Raw BSG Report")

        # Create Report with Raw BSG Report
        workbook = Workbook()
        del workbook["Sheet"]

        for jobStep in jobs:
            jobStep.reportData(workbook, controllerData, type(jobStep).__name__, False, False)

        logging.debug(f"Saving Raw BSG Report Workbook")
        workbook.save(f"output/{jobFileName}/{jobFileName}-RawBSGReport.xlsx")
