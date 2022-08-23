import logging
from datetime import datetime

from openpyxl import Workbook
from output.ReportBase import ReportBase
from util.excel_utils import Color, addFilterAndFreeze, resizeColumnWidth, writeRow, writeUncoloredRow


class DashboardReport(ReportBase):
    def createWorkbook(self, jobs, controllerData, jobFileName):
        logging.info(f"Creating Dashboard Report Workbook")

        workbook = Workbook()
        del workbook["Sheet"]

        logging.debug(f"Creating workbook sheet for Dashboards")
        dashboardSheet = workbook.create_sheet(f"Dashboards")
        writeUncoloredRow(dashboardSheet, 1, ["controller", "dashboardName"])

        # Write Data
        rowIdx = 2
        for host, hostInfo in controllerData.items():
            for dashboard in hostInfo["exportedDashboards"]:
                color = None

                writeRow(
                    dashboardSheet,
                    rowIdx,
                    [
                        (hostInfo["controller"].host, color),
                        (dashboard["name"], color),

                    ],
                )
                rowIdx += 1

        addFilterAndFreeze(dashboardSheet)
        resizeColumnWidth(dashboardSheet)

        logging.debug(f"Saving Dashboard Workbook")
        workbook.save(f"output/{jobFileName}/{jobFileName}-Dashboards.xlsx")
