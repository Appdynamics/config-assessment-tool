import logging
from datetime import datetime

from openpyxl import Workbook
from reports.ReportBase import ReportBase
from util.xcel_utils import addFilterAndFreeze, resizeColumnWidth, writeColoredRow, writeSummarySheet, writeUncoloredRow


class SyntheticsReport(ReportBase):
    def createWorkbook(self, jobs, controllerData, jobFileName):
        logging.info(f"Creating Synthetics Report Workbook")

        # Create Report with Raw Data
        workbook = Workbook()

        summarySheet = workbook["Sheet"]
        summarySheet.title = "Synthetics"

        allSyntheticJobs = []
        for host, hostInfo in controllerData.items():
            for application in hostInfo["brum"].values():
                for syntheticJob in application["syntheticJobs"]["jobListDatas"]:
                    billableTimeAverage24Hr = syntheticJob["billableTime"]["billableTimeAverage24Hr"] if syntheticJob["billableTime"] else None
                    currentMonthBillableTimeTotal = (
                        syntheticJob["billableTime"]["currentMonthBillableTimeTotal"] if syntheticJob["billableTime"] else None
                    )
                    privateAgentUtilization = (
                        syntheticJob["privateAgentUtilization"]["utilization"] if syntheticJob["privateAgentUtilization"] else None
                    )
                    allSyntheticJobs.append(
                        {
                            "host": host,
                            "componentType": "brum",
                            "application": application["name"],
                            "jobName": syntheticJob["config"]["description"],
                            "projectedDailyRuns": syntheticJob["config"]["projectedUsage"]["projectedDailyRuns"],
                            "projectedMonthlyRuns": syntheticJob["config"]["projectedUsage"]["projectedMonthlyRuns"],
                            "billableTimeAverage24Hr": billableTimeAverage24Hr,
                            "currentMonthBillableTimeTotal": currentMonthBillableTimeTotal,
                            "privateAgentUtilization": privateAgentUtilization,
                            "browsers": ",".join(syntheticJob["config"]["browserCodes"]),
                            "usesPrivateAgent": syntheticJob["hasPrivateAgent"],
                            "created": datetime.fromtimestamp(syntheticJob["config"]["created"] / 1000.0),
                            "updated": datetime.fromtimestamp(syntheticJob["config"]["updated"] / 1000.0),
                        }
                    )

        if len(allSyntheticJobs) == 0:
            logging.warning(f"No data found for Synthetics")
            return

        # Write Headers
        writeUncoloredRow(
            summarySheet,
            1,
            [
                *allSyntheticJobs[0].keys(),
            ],
        )

        rowIdx = 2
        for syntheticJob in allSyntheticJobs:
            writeUncoloredRow(
                summarySheet,
                rowIdx,
                [*syntheticJob.values()],
            )
            rowIdx += 1

        addFilterAndFreeze(summarySheet, freezePane="E2")
        resizeColumnWidth(summarySheet)

        logging.debug(f"Saving Synthetics Workbook")
        workbook.save(f"output/{jobFileName}/{jobFileName}-Synthetics.xlsx")
