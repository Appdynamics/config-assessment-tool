import logging
import os.path

import pandas as pd
import xlsxwriter

from output.PostProcessReport import PostProcessReport

light_font = '#FFFFFF'
dark_font = '#000000'

medium_bg = '#A0A0A0'
dark_bg = '#000000'


class ConfigurationAnalysisReport(PostProcessReport):

    def __init__(self):
        self.workbook = None
        self.analysis_sheet = None

    async def post_process(self, jobFileName):

        logging.info(f"Running post-process command: Creating Configuration Analysis Workbook")

        directory = f"output/{jobFileName}"
        file_prefix = f"{jobFileName}"
        # input
        self.analysis_sheet = os.path.join(directory, f"{file_prefix}-MaturityAssessment-apm.xlsx")

        if not os.path.exists(self.analysis_sheet):
            logging.warn(f"Input file sheet {self.analysis_sheet} does not exist. "
                         f"Skipping post-process command: Creating Configuration Analysis Workbook")
            return

        # output
        self.workbook = xlsxwriter.Workbook(f"output/{jobFileName}/{jobFileName}-ConfigurationAnalysisReport.xlsx")
        worksheets = self.generateHeaders()
        applicationNames = self.getListOfApplications()
        applicationData = []

        for application in applicationNames:
            taskList = [[], [], [], [], []]
            ranking = self.performAnalysis(application, taskList)
            if ranking != "Platinum":
                for task in taskList:
                    if (len(task) > 0):
                        applicationData.append([application, ranking, taskList])
                        break
            else:
                applicationData.append([application, ranking, []])

        self.buildOutput(applicationData, worksheets)
        logging.info(f"Saving ConfigurationAnalysisReport Workbook")
        self.workbook.close()

    def getValuesInColumn(self, sheet, col1_value):
        values = []
        for column_cell in sheet.iter_cols(1, sheet.max_column):
            if column_cell[0].value == col1_value:
                j = 0
                for data in column_cell[1:]:
                    values.append(data.value)
                break
        return values

    def getListOfApplications(self):
        frame = pd.read_excel(self.analysis_sheet, sheet_name='Analysis',
                              engine='openpyxl').dropna(how='all')
        return frame['name'].tolist()

    def overallAppStatus(self, application, tasklist):
        frame = pd.read_excel(self.analysis_sheet, sheet_name='Analysis', engine='openpyxl')
        frame = frame.drop('controller', axis=1)
        appFrame = frame.loc[frame['name'] == application]

        # Overall Assessment
        ranking = 'NA'
        if (appFrame['OverallAssessment'] == 'bronze').any():
            ranking = 'Bronze'
        elif (appFrame['OverallAssessment'] == 'silver').any():
            ranking = 'Silver'
        elif (appFrame['OverallAssessment'] == 'gold').any():
            ranking = 'Gold'
        elif (appFrame['OverallAssessment'] == 'platinum').any():
            ranking = 'Platinum'
        return ranking

    def appAgentStatus(self, application, taskList):
        # Sheet name may have changed to AppAgentsAPM
        frame = pd.read_excel(self.analysis_sheet, sheet_name='AppAgentsAPM', engine='openpyxl')
        frame = frame.drop('controller', axis=1)
        appFrame = frame.loc[frame['application'] == application]

        # Agent Metric Limit
        if (appFrame['metricLimitNotHit'] == False).any():
            taskList[2].append("Application Agent metric limit has been reached")

        # Agent Versions
        if (appFrame['percentAgentsLessThan2YearsOld'] < 50).any():
            taskList[0].append(str(100 - int(appFrame['percentAgentsLessThan2YearsOld'])) + '% of Application Agents are 2+ years old')
        elif (appFrame['percentAgentsLessThan1YearOld'] < 80).any():
            taskList[0].append(str(100 - int(appFrame['percentAgentsLessThan1YearOld'])) + '% of Application Agents are at least 1 year old')

        # Agents reporting data
        if (appFrame['percentAgentsReportingData'] < 100).any():
            taskList[0].append(str(100 - int(appFrame['percentAgentsReportingData'])) + "% of Application Agents aren't reporting data")

        if (appFrame['percentAgentsRunningSameVersion'] < 100).any():
            taskList[0].append('Multiple Application Agent Versions')

    def machineAgentStatus(self, application, taskList):
        frame = pd.read_excel(self.analysis_sheet, sheet_name='MachineAgentsAPM', engine='openpyxl')
        frame = frame.drop('controller', axis=1)
        appFrame = frame.loc[frame['application'] == application]

    def businessTranStatus(self, application, taskList):
        frame = pd.read_excel(self.analysis_sheet, sheet_name='BusinessTransactionsAPM', engine='openpyxl')
        frame.drop('controller', axis=1)
        appFrame = frame.loc[frame['application'] == application]

        # Number of Business Transcations
        if (appFrame['numberOfBTs'] > 200).any():
            taskList[1].append("Reduce amount of Business transactions from " + str(int(appFrame['numberOfBTs'])))

        # % of Business Transactions with load
        if (appFrame['percentBTsWithLoad'] < 90).any():
            taskList[1].append(str(100 - int(appFrame['percentBTsWithLoad'])) + '% of Business Transactions have no load over the last 24 hours')

        # Business Transaction Lockdown
        if (appFrame['btLockdownEnabled'] == False).any():
            taskList[1].append("Business Transaction Lockdown is disabled")

        # Number of Custom Match Rules
        if (appFrame['numberCustomMatchRules'] < 3).any():
            if (appFrame['numberCustomMatchRules'] == 0).any():
                taskList[2].append('No Custom Match Rules')
            else:
                taskList[2].append('Only ' + str(int(appFrame['numberCustomMatchRules'])) + ' Custom Match Rules')

    def backendStatus(self, application, taskList):
        frame = pd.read_excel(self.analysis_sheet, sheet_name='BackendsAPM', engine='openpyxl')
        frame.drop('controller', axis=1)
        appFrame = frame.loc[frame['application'] == application]

        # % of Backends with load
        if (appFrame['percentBackendsWithLoad'] < 75).any():
            taskList[2].append(str(100 - int(appFrame['percentBackendsWithLoad'])) + '% of Backends have no load')

        # Backend limit not hit
        if (appFrame['backendLimitNotHit'] == False).any():
            taskList[2].append('Backend limit has been reached')

        # Number of Custom Backend Rules
        if (appFrame['numberOfCustomBackendRules'] == 0).any():
            taskList[2].append('No Custom Backend Rules')

    def overheadStatus(self, application, taskList):
        frame = pd.read_excel(self.analysis_sheet, sheet_name='OverheadAPM', engine='openpyxl')
        frame.drop('controller', axis=1)
        appFrame = frame.loc[frame['application'] == application]

        # Developer Mode Not Enabled for any Business Transaction
        if (appFrame['developerModeNotEnabledForAnyBT'] == False).any():
            taskList[2].append('Development Level monitoring is enabled for a Business Transaction')

        # find-entry-points not enabled
        if (appFrame['findEntryPointsNotEnabled'] == False).any():
            taskList[2].append('Find-entry-points node property is enabled')

        # Aggressive Snapshotting not enabled
        if (appFrame['aggressiveSnapshottingNotEnabled'] == False).any():
            taskList[2].append('Aggressive snapshot collection is enabled')

        # Developer Mode not enabled for an application
        if (appFrame['developerModeNotEnabledForApplication'] == False).any():
            taskList[2].append('Development Level monitoring is enabled for an Application')

    def serviceEndpointStatus(self, application, taskList):
        frame = pd.read_excel(self.analysis_sheet, sheet_name='ServiceEndpointsAPM', engine='openpyxl')
        frame.drop('controller', axis=1)
        appFrame = frame.loc[frame['application'] == application]

        # Number of Custom Service Endpoint Rules
        if (appFrame['numberOfCustomServiceEndpointRules'] == 0).any():
            taskList[2].append('No Custom Service Endpoint rules')

        # Service Endpoint Limit not hit
        if (appFrame['serviceEndpointLimitNotHit'] == False).any():
            taskList[2].append('Service Endpoint limit has been reached')

        # % of enabled Service Endpoints with load
        if (appFrame['percentServiceEndpointsWithLoadOrDisabled'] < 75).any():
            taskList[2].append(str(100 - int(appFrame['percentServiceEndpointsWithLoadOrDisabled'])) + '% of enabled Service Endpoints have no load')

    def errorConfigurationStatus(self, application, taskList):
        frame = pd.read_excel(self.analysis_sheet, sheet_name='ErrorConfigurationAPM', engine='openpyxl')
        frame.drop('controller', axis=1)
        appFrame = frame.loc[frame['application'] == application]

        # Sucess Percentage of Worst Transaction
        if (appFrame['successPercentageOfWorstTransaction'] < 80).any():
            taskList[3].append('Some Business Transactions fail ' + str(100 - int(appFrame['successPercentageOfWorstTransaction'])) + '% of the time')

        # Number of Custom rules
        if (appFrame['numberOfCustomRules'] == 0).any():
            taskList[2].append('No custom error configurations')

    def healthRulesAlertingStatus(self, application, taskList):
        frame = pd.read_excel(self.analysis_sheet, sheet_name='HealthRulesAndAlertingAPM', engine='openpyxl')
        frame.drop('controller', axis=1)
        appFrame = frame.loc[frame['application'] == application]

        # Number of Health Rule Violations in last 24 hours
        if (appFrame['numberOfHealthRuleViolations'] > 10).any():
            taskList[3].append(str(int(appFrame['numberOfHealthRuleViolations'])) + ' Health Rule Violations in 24 hours')

        # Number of modifications to default Health Rules
        if (appFrame['numberOfDefaultHealthRulesModified'] < 2).any():
            if (appFrame['numberOfDefaultHealthRulesModified'] < 2).any():
                taskList[3].append('No modifications to the default Health Rules')
            else:
                taskList[3].append('Only ' + str(int(appFrame['numberOfDefaultHealthRulesModified'])) + ' modifications to the default Health Rules')

        # Number of actions bound to enabled policies
        if (appFrame['numberOfActionsBoundToEnabledPolicies'] < 1).any():
            taskList[3].append('No actions bound to enabled policies')

        # Number of Custom Health Rules
        if (appFrame['numberOfCustomHealthRules'] < 5).any():
            if (appFrame['numberOfCustomHealthRules'] == 0).any():
                taskList[3].append('No Custom Health Rules')
            else:
                taskList[3].append('Only ' + str(int(appFrame['numberOfCustomHealthRules'])) + ' Custom Health Rules')

    def dataCollectorStatus(self, application, taskList):
        frame = pd.read_excel(self.analysis_sheet, sheet_name='DataCollectorsAPM', engine='openpyxl')
        frame.drop('controller', axis=1)
        appFrame = frame.loc[frame['application'] == application]

        # Number of data collector fields configured
        if (appFrame['numberOfDataCollectorFieldsConfigured'] < 5).any():
            if (appFrame['numberOfDataCollectorFieldsConfigured'] == 0).any():
                taskList[2].append('No configured Data Collectors')
            else:
                taskList[2].append('Only ' + str(int(appFrame['numberOfDataCollectorFieldsConfigured'])) + ' configured Data Collectors')

        # Number of data collector fields colleced in snapshots in last 24 hours
        if (appFrame['numberOfDataCollectorFieldsCollectedInSnapshots'] < 5).any():
            if (appFrame['numberOfDataCollectorFieldsCollectedInSnapshots'] == 0).any():
                taskList[2].append('No Data Collector fields collected in APM Snapshots in 24 hours')
            else:
                taskList[2].append('Only ' + str(int(appFrame['numberOfDataCollectorFieldsCollectedInSnapshots'])) + ' Data Collector fields collected in APM Snapshots in 24 hours')

        # Number of data collector fields collect in analytics in last 24 hours
        if (appFrame['numberOfDataCollectorFieldsCollectedInAnalytics'] < 5).any():
            if (appFrame['numberOfDataCollectorFieldsCollectedInAnalytics'] == 0).any():
                taskList[2].append('No Data Collector fields collected in Analytics in 24 hours')
            else:
                taskList[2].append('Only ' + str(int(appFrame['numberOfDataCollectorFieldsCollectedInAnalytics'])) + ' Data Collector fields collected in Analytics in 24 hours')

        # BiQ enabled
        if (appFrame['biqEnabled'] == False).any():
            taskList[2].append('BiQ is disabled')

    def apmDashBoardsStatus(self, application, taskList):
        frame = pd.read_excel(self.analysis_sheet, sheet_name='DashboardsAPM', engine='openpyxl')
        frame.drop('controller', axis=1)
        appFrame = frame.loc[frame['application'] == application]

        # Number of custom dashboards
        if (appFrame['numberOfDashboards'] < 5).any():
            if (appFrame['numberOfDashboards'] == 1).any():
                taskList[4].append('Only 1 Custom Dashboard')
            elif (appFrame['numberOfDashboards'] == 0).any():
                taskList[4].append('No Custom Dashboards')
            else:
                taskList[4].append('Only ' + str(int(appFrame['numberOfDashboards'])) + ' Custom Dashboards')

        # % of Custom Dashboards modified in last 6 months
        if (appFrame['percentageOfDashboardsModifiedLast6Months'] < 100).any():
            taskList[4].append(str(100 - int(appFrame['percentageOfDashboardsModifiedLast6Months'])) + '% of Custom Dashboards have not been updated in 6+ months')

        # Number of Custom Dashboards using BiQ
        if (appFrame['numberOfDashboardsUsingBiQ'] == 0).any():
            taskList[4].append('No Custom Dashboards using BiQ')

    def performAnalysis(self, application, taskList):
        overallRanking = self.overallAppStatus(application, taskList)
        self.appAgentStatus(application, taskList)
        self.machineAgentStatus(application, taskList)
        self.businessTranStatus(application, taskList)
        self.backendStatus(application, taskList)
        self.overheadStatus(application, taskList)
        self.serviceEndpointStatus(application, taskList)
        self.errorConfigurationStatus(application, taskList)
        self.healthRulesAlertingStatus(application, taskList)
        self.dataCollectorStatus(application, taskList)
        self.apmDashBoardsStatus(application, taskList)

        return overallRanking

    def buildOutput(self, applicationData, worksheets):
        worksheet = None

        stepFormat = self.workbook.add_format()
        stepFormat.set_align('center')
        # row_num = 1
        row_counts = [1, 1, 1, 1]

        for application in range(0, len(applicationData)):
            applicationName = applicationData[application][0]
            applicationRank = applicationData[application][1]
            currentApplication = applicationData[application][2]

            if applicationRank == 'Bronze':
                worksheet = worksheets[0]
                self.generateApplicationHeader(applicationName, applicationRank, worksheet, row_counts[0])
                row_counts[0] = row_counts[0] + 1
            elif applicationRank == 'Silver':
                worksheet = worksheets[1]
                self.generateApplicationHeader(applicationName, applicationRank, worksheet, row_counts[1])
                row_counts[1] = row_counts[1] + 1
            elif applicationRank == 'Gold':
                worksheet = worksheets[2]
                self.generateApplicationHeader(applicationName, applicationRank, worksheet, row_counts[2])
                row_counts[2] = row_counts[2] + 1
            elif applicationRank == 'Platinum':
                worksheet = worksheets[3]
                self.generateApplicationHeader(applicationName, applicationRank, worksheet, row_counts[3])
                row_counts[3] = row_counts[3] + 1

            # row_num += 1
            task_num = 1
            for categoryIndex in range(0, len(currentApplication)):
                for taskIndex in range(0, len(currentApplication[categoryIndex])):

                    if applicationRank == 'Bronze':
                        worksheet.write(row_counts[0], 0, applicationName)
                        worksheet.write(row_counts[0], 1, task_num, stepFormat)
                        worksheet.write(row_counts[0], 3, currentApplication[categoryIndex][taskIndex])
                        # worksheet.write(row_num, 4, applicationRank)
                        worksheet.write(row_counts[0], 4, " ")

                        if categoryIndex == 0:
                            worksheet.write(row_counts[0], 2, "Agent Installation (APM, Machine, Server, etc.)")
                        elif categoryIndex == 1:
                            worksheet.write(row_counts[0], 2, "Business Transactions")
                        elif categoryIndex == 2:
                            worksheet.write(row_counts[0], 2, "Advanced Configurations")
                        elif categoryIndex == 3:
                            worksheet.write(row_counts[0], 2, "Health Rules & Alerts")
                        elif categoryIndex == 4:
                            worksheet.write(row_counts[0], 2, "Dashboard")
                        worksheet.set_row(row_counts[0], None, None, {'level': 1})
                        row_counts[0] += 1
                        task_num += 1
                    elif applicationRank == 'Silver':
                        worksheet.write(row_counts[1], 0, applicationName)
                        worksheet.write(row_counts[1], 1, task_num, stepFormat)
                        worksheet.write(row_counts[1], 3, currentApplication[categoryIndex][taskIndex])
                        # worksheet.write(row_num, 4, applicationRank)
                        worksheet.write(row_counts[1], 4, " ")

                        if categoryIndex == 0:
                            worksheet.write(row_counts[1], 2, "Agent Installation (APM, Machine, Server, etc.)")
                        elif categoryIndex == 1:
                            worksheet.write(row_counts[1], 2, "Business Transactions")
                        elif categoryIndex == 2:
                            worksheet.write(row_counts[1], 2, "Advanced Configurations")
                        elif categoryIndex == 3:
                            worksheet.write(row_counts[1], 2, "Health Rules & Alerts")
                        elif categoryIndex == 4:
                            worksheet.write(row_counts[1], 2, "Dashboard")
                        worksheet.set_row(row_counts[1], None, None, {'level': 1})
                        row_counts[1] += 1
                        task_num += 1
                    elif applicationRank == 'Gold':
                        worksheet.write(row_counts[2], 0, applicationName)
                        worksheet.write(row_counts[2], 1, task_num, stepFormat)
                        worksheet.write(row_counts[2], 3, currentApplication[categoryIndex][taskIndex])
                        # worksheet.write(row_num, 4, applicationRank)
                        worksheet.write(row_counts[2], 4, " ")

                        if categoryIndex == 0:
                            worksheet.write(row_counts[2], 2, "Agent Installation (APM, Machine, Server, etc.)")
                        elif categoryIndex == 1:
                            worksheet.write(row_counts[2], 2, "Business Transactions")
                        elif categoryIndex == 2:
                            worksheet.write(row_counts[2], 2, "Advanced Configurations")
                        elif categoryIndex == 3:
                            worksheet.write(row_counts[2], 2, "Health Rules & Alerts")
                        elif categoryIndex == 4:
                            worksheet.write(row_counts[2], 2, "Dashboard")
                        worksheet.set_row(row_counts[2], None, None, {'level': 1})
                        row_counts[2] += 1
                        task_num += 1
                    elif applicationRank == 'Platinum':
                        worksheet.write(row_counts[3], 0, applicationName)
                        worksheet.write(row_counts[3], 1, task_num, stepFormat)
                        worksheet.write(row_counts[3], 3, currentApplication[categoryIndex][taskIndex])
                        # worksheet.write(row_num, 4, applicationRank)
                        worksheet.write(row_counts[3], 4, " ")

                        if categoryIndex == 0:
                            worksheet.write(row_counts[3], 2, "Agent Installation (APM, Machine, Server, etc.)")
                        elif categoryIndex == 1:
                            worksheet.write(row_counts[3], 2, "Business Transactions")
                        elif categoryIndex == 2:
                            worksheet.write(row_counts[3], 2, "Advanced Configurations")
                        elif categoryIndex == 3:
                            worksheet.write(row_counts[3], 2, "Health Rules & Alerts")
                        elif categoryIndex == 4:
                            worksheet.write(row_counts[3], 2, "Dashboard")
                        worksheet.set_row(row_counts[3], None, None, {'level': 1})
                        row_counts[3] += 1
                        task_num += 1

        for i in range(len(worksheets)):
            worksheets[i].autofilter('A1:D' + str(row_counts[i]))

    def generateApplicationHeader(self, applicationName, applicationRank, worksheet, row_num):
        appHeaderFormat = self.workbook.add_format()
        appHeaderFormat.set_bold()
        appHeaderFormat.set_font_color(dark_font)
        appHeaderFormat.set_bg_color(medium_bg)

        stepsHeaderFormat = self.workbook.add_format()
        stepsHeaderFormat.set_bold()
        stepsHeaderFormat.set_font_color(dark_font)
        stepsHeaderFormat.set_bg_color(medium_bg)
        stepsHeaderFormat.set_align('center')

        worksheet.write(row_num, 0, applicationName, appHeaderFormat)
        worksheet.write(row_num, 1, " ", stepsHeaderFormat)
        worksheet.write(row_num, 2, " ", appHeaderFormat)
        worksheet.write(row_num, 3, " ", appHeaderFormat)
        # worksheet.write(row_num, 4, applicationRank, appHeaderFormat)
        worksheet.write(row_num, 4, " ", appHeaderFormat)

    def generateHeaders(self):
        bronze_worksheet = self.workbook.add_worksheet('Bronze')
        silver_worksheet = self.workbook.add_worksheet('Silver')
        gold_worksheet = self.workbook.add_worksheet('Gold')
        plat_worksheet = self.workbook.add_worksheet('Platinum')

        worksheets = [bronze_worksheet, silver_worksheet, gold_worksheet, plat_worksheet]

        headerFormat = self.workbook.add_format()
        headerFormat.set_bold()
        headerFormat.set_font_color(light_font)
        headerFormat.set_bg_color(dark_bg)

        stepFormat = self.workbook.add_format()
        stepFormat.set_bold()
        stepFormat.set_font_color(light_font)
        stepFormat.set_bg_color(dark_bg)
        stepFormat.set_align('center')

        for sheet in range(len(worksheets)):
            worksheets[sheet].write('A1', 'Application', headerFormat)
            worksheets[sheet].write('B1', 'Steps', stepFormat)
            worksheets[sheet].write('C1', 'Activity', headerFormat)
            worksheets[sheet].write('D1', 'Task', headerFormat)
            worksheets[sheet].write('E1', 'Target', headerFormat)

            # worksheet.write('E1', 'Overall Ranking', headerFormat)

            worksheets[sheet].set_column('A:A', 40)
            worksheets[sheet].set_column('B:B', 7)
            worksheets[sheet].set_column('C:C', 50)
            worksheets[sheet].set_column('D:D', 65)

            # worksheet.set_column('E:E', 15)

            worksheets[sheet].freeze_panes(1, 0)

        return worksheets
