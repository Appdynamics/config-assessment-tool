import logging

from openpyxl import Workbook
from output.ReportBase import ReportBase
from util.excel_utils import addFilterAndFreeze, resizeColumnWidth, writeColoredRow, writeSummarySheet, writeUncoloredRow


class CustomMetricsReport(ReportBase):
    def createWorkbook(self, jobs, controllerData, jobFileName):
        logging.info(f"Creating Custom Metrics Workbook")

        # Create Report with Raw Data
        workbook = Workbook()

        summarySheet = workbook["Sheet"]
        summarySheet.title = "Extensions"

        allExtensions = set()
        for host, hostInfo in controllerData.items():
            for application in hostInfo["apm"].values():
                for extension in application["customMetrics"]:
                    allExtensions.add(extension)

        # Write Headers
        writeUncoloredRow(
            summarySheet,
            1,
            [
                "controller",
                "componentType",
                "application",
                *allExtensions,
            ],
        )

        rowIdx = 2
        for host, hostInfo in controllerData.items():
            for component in hostInfo["apm"].values():
                writeUncoloredRow(
                    summarySheet,
                    rowIdx,
                    [
                        hostInfo["controller"].host,
                        "apm",
                        component["name"],
                        *[(True if e in component["customMetrics"] else False) for e in allExtensions],
                    ],
                )
                rowIdx += 1

        addFilterAndFreeze(summarySheet)
        resizeColumnWidth(summarySheet)

        logging.debug(f"Saving CustomMetrics Workbook")
        workbook.save(f"output/{jobFileName}/{jobFileName}-CustomMetrics.xlsx")
