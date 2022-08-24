import logging
from datetime import datetime
from math import ceil, floor

from openpyxl import Workbook
from output.ReportBase import ReportBase
from util.excel_utils import Color, addFilterAndFreeze, resizeColumnWidth, writeColoredRow, writeSummarySheet, writeUncoloredRow


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
                    averageBlocksUsedPerRun = ceil(floor(syntheticJob["averageDuration"]) / 1000 / 5) if not syntheticJob["hasPrivateAgent"] else 0
                    estimatedBlocksUsedPerMonthPercentage = (
                        averageBlocksUsedPerRun
                        * syntheticJob["config"]["projectedUsage"]["projectedMonthlyRuns"]
                        / hostInfo["eumLicenseUsage"]["allocatedSyntheticMeasurementUnits"]
                        * 100
                    )
                    color = Color.white
                    if estimatedBlocksUsedPerMonthPercentage > 100:
                        color = Color.red
                    allSyntheticJobs.append(
                        {
                            "host": (host, color),
                            "componentType": ("brum", color),
                            "application": (application["name"], color),
                            "jobName": (syntheticJob["config"]["description"], color),
                            "projectedDailyRuns": (syntheticJob["config"]["projectedUsage"]["projectedDailyRuns"], color),
                            "projectedMonthlyRuns": (syntheticJob["config"]["projectedUsage"]["projectedMonthlyRuns"], color),
                            "averageDuration": (syntheticJob["averageDuration"], color),
                            "averageBlocksUsedPerRun": (averageBlocksUsedPerRun, color),
                            "estimatedBlocksUsedPerMonth": (
                                averageBlocksUsedPerRun * syntheticJob["config"]["projectedUsage"]["projectedMonthlyRuns"],
                                color,
                            ),
                            "estimatedBlocksUsedPerMonthPercentage": (estimatedBlocksUsedPerMonthPercentage, color),
                            "billableTimeAverage24Hr": (billableTimeAverage24Hr, color),
                            "currentMonthBillableTimeTotal": (currentMonthBillableTimeTotal, color),
                            "privateAgentUtilization": (privateAgentUtilization, color),
                            "browsers": (",".join(syntheticJob["config"]["browserCodes"]), color),
                            "usesPrivateAgent": (syntheticJob["hasPrivateAgent"], color),
                            "created": (datetime.fromtimestamp(syntheticJob["config"]["created"] / 1000.0), color),
                            "updated": (datetime.fromtimestamp(syntheticJob["config"]["updated"] / 1000.0), color),
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
            writeColoredRow(
                summarySheet,
                rowIdx,
                [*syntheticJob.values()],
            )
            rowIdx += 1

        addFilterAndFreeze(summarySheet, freezePane="E2")
        resizeColumnWidth(summarySheet)

        logging.debug(f"Saving Synthetics Workbook")
        workbook.save(f"output/{jobFileName}/{jobFileName}-Synthetics.xlsx")
