import logging

from openpyxl import Workbook

from util.xcel_utils import (
    writeUncoloredRow,
    writeColoredRow,
    addFilterAndFreeze,
    resizeColumnWidth,
    writeSummarySheet,
)

from reports.ReportBase import ReportBase


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
                        *[(True if e in component['customMetrics'] else False) for e in allExtensions],
                    ],
                )
                rowIdx += 1

        addFilterAndFreeze(summarySheet)
        resizeColumnWidth(summarySheet)

        logging.debug(f"Saving Custom Metrics Workbook")
        workbook.save(f"output/{jobFileName}/{jobFileName}-CustomMetricsReport.xlsx")
