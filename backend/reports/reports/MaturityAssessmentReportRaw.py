import logging

from openpyxl import Workbook
from reports.ReportBase import ReportBase


class RawMaturityAssessmentReport(ReportBase):
    def createWorkbook(self, jobs, controllerData, jobFileName):
        for reportType in ["apm", "brum", "mrum"]:
            logging.info(f"Creating {reportType} Maturity Assessment Raw Report")

            # Create Report with Raw Maturity Assessment Report
            workbook = Workbook()

            filteredJobs = [job for job in jobs if job.componentType == reportType]

            if filteredJobs:
                del workbook["Sheet"]
            else:
                summarySheet = workbook["Sheet"]
                summarySheet.title = "Summary"

            for jobStep in filteredJobs:
                jobStep.reportData(workbook, controllerData, type(jobStep).__name__, False, False)

            logging.debug(f"Saving Raw MaturityAssessment-{reportType} Workbook")
            workbook.save(f"output/{jobFileName}/{jobFileName}-MaturityAssessmentRaw-{reportType}.xlsx")
