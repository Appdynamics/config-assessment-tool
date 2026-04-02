import logging
import os

from openpyxl import Workbook
from backend.output.ReportBase import ReportBase
from backend.util.excel_utils import addFilterAndFreeze, resizeColumnWidth, writeColoredRow, writeSummarySheet, writeUncoloredRow


class CustomMetricsReport(ReportBase):
    def createWorkbook(self, jobs, controllerData, jobFileName, output_dir="output"):
        logging.info(f"Creating Custom Metrics Report Workbook")

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
                "applicationId",
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
                        component["id"],
                        *[(True if e in component["customMetrics"] else False) for e in allExtensions],
                    ],
                )
                rowIdx += 1

        addFilterAndFreeze(summarySheet, "E2")
        resizeColumnWidth(summarySheet)

        logging.debug(f"Saving CustomMetrics Workbook")
        save_path = os.path.join(output_dir, jobFileName, f"{jobFileName}-CustomMetrics.xlsx")
        workbook.save(save_path)
