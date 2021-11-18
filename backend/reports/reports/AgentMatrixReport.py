import logging
from collections import Counter

from openpyxl import Workbook

from util.xcel_utils import writeUncoloredRow, addFilterAndFreeze, resizeColumnWidth

from reports.ReportBase import ReportBase


class AgentMatrixReport(ReportBase):
    def createWorkbook(self, jobs, controllerData, jobFileName):
        logging.info(f"Creating Agent Matrix Report Workbook")

        workbook = Workbook()
        del workbook["Sheet"]

        allAppAgentVersions = set()
        for host in controllerData.values():
            allAppAgentVersions.update(host["appAgentVersions"])
        allAppAgentVersions = [f"{version[2]}:{version[0]}.{version[1]}" for version in sorted(list(allAppAgentVersions), reverse=True)]

        logging.debug(f"Creating workbook sheet for App Agents")
        appAgentsSheet = workbook.create_sheet(f"App Agents")
        writeUncoloredRow(appAgentsSheet, 1, ["controller", "application", *allAppAgentVersions])
        # Write Data
        rowIdx = 2
        for host, hostInfo in controllerData.items():
            for application in hostInfo["apm"].values():
                agentVersionMap = Counter(application["appAgentVersions"])
                agentCountRow = []
                for version in allAppAgentVersions:
                    if version in agentVersionMap:
                        agentCountRow.append(agentVersionMap[version])
                    else:
                        agentCountRow.append(0)

                writeUncoloredRow(
                    appAgentsSheet,
                    rowIdx,
                    [
                        hostInfo["controller"].host,
                        application["name"],
                        *agentCountRow,
                    ],
                )
                rowIdx += 1

        addFilterAndFreeze(appAgentsSheet)
        resizeColumnWidth(appAgentsSheet)

        logging.debug(f"Creating workbook sheet for Machine Agents")

        allMachineAgentVersions = set()
        for host in controllerData.values():
            allMachineAgentVersions.update(host["machineAgentVersions"])
        allMachineAgentVersions = [f"{version[0]}.{version[1]}" for version in sorted(list(allMachineAgentVersions), reverse=True)]

        machineAgentsSheet = workbook.create_sheet(f"Machine Agents")
        writeUncoloredRow(
            machineAgentsSheet,
            1,
            ["controller", "application", *allMachineAgentVersions],
        )
        # Write Data
        rowIdx = 2
        for host, hostInfo in controllerData.items():
            for application in hostInfo["apm"].values():
                agentVersionMap = Counter(application["machineAgentVersions"])
                agentCountRow = []
                for version in allMachineAgentVersions:
                    if version in agentVersionMap:
                        agentCountRow.append(agentVersionMap[version])
                    else:
                        agentCountRow.append(0)

                writeUncoloredRow(
                    machineAgentsSheet,
                    rowIdx,
                    [
                        hostInfo["controller"].host,
                        application["name"],
                        *agentCountRow,
                    ],
                )
                rowIdx += 1

        addFilterAndFreeze(machineAgentsSheet)
        resizeColumnWidth(machineAgentsSheet)

        logging.debug(f"Saving Agent Matrix Workbook")
        workbook.save(f"output/{jobFileName}/{jobFileName}-Agent-Matrix.xlsx")
