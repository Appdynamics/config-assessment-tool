import os
import pytest
import output.presentations.cxPptFsoUseCases as fsoppt
from pptx import Presentation
from output.presentations.cxPptFsoUseCases import UseCase, ExcelSheets

def search(filename, search_path="."):
    for root, _, files in os.walk(search_path):
        if filename in files:
            return os.path.join(root, filename)
    return None

@pytest.fixture
def uc():
    json_file = search("HybridApplicationMonitoringUseCase.json", "../")
    return UseCase(json_file)


@pytest.fixture
def excels():
    folder = f"output/test_ham_report"
    ma_apm = 'test_ham_report-MaturityAssessment-apm.xlsx'
    ma_apm_raw = 'test_ham_report-MaturityAssessmentRaw-apm.xlsx'
    agent_matrix = 'test_ham_report-AgentMatrix.xlsx'
    return ExcelSheets(folder,(ma_apm,ma_apm_raw, agent_matrix))



def testCxPptFsoUseCases(uc,excels):
    template = search("HybridApplicationMonitoringUseCase_template.pptx", "../")
    root = Presentation(template)
    data = uc.pitstop_data
    setHealthStatus(uc)

    fsoppt.markRaceTrackFailures(root,uc)

    fsoppt.generatePitstopHealthCheckTable("folder_xxx", root, uc, "onboard")
    fsoppt.generatePitstopHealthCheckTable("folder_xxx", root, uc, "implement")
    fsoppt.generatePitstopHealthCheckTable("folder_xxx", root, uc, "use")
    fsoppt.generatePitstopHealthCheckTable("folder_xxx", root, uc, "engage")
    fsoppt.generatePitstopHealthCheckTable("folder_xxx", root, uc, "adopt")
    fsoppt.generatePitstopHealthCheckTable("folder_xxx", root, uc, "optimize")

    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "onboard", fsoppt.SlideId.ONBOARD_REMEDIATION_PRIMARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "implement", fsoppt.SlideId.IMPLEMENT_REMEDIATION_PRIMARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "engage", fsoppt.SlideId.ENGAGE_REMEDIATION_PRIMARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "adopt", fsoppt.SlideId.ADOPT_REMEDIATION_PRIMARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "optimize", fsoppt.SlideId.OPTIMIZE_REMEDIATION_PRIMARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "use", fsoppt.SlideId.USE_REMEDIATION_PRIMARY)

    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "onboard", fsoppt.SlideId.ONBOARD_REMEDIATION_SECONDARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "implement", fsoppt.SlideId.IMPLEMENT_REMEDIATION_SECONDARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "engage", fsoppt.SlideId.ENGAGE_REMEDIATION_SECONDARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "adopt", fsoppt.SlideId.ADOPT_REMEDIATION_SECONDARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "optimize", fsoppt.SlideId.OPTIMIZE_REMEDIATION_SECONDARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "use", fsoppt.SlideId.USE_REMEDIATION_SECONDARY)

    root.save(f"cx-ham-usecase-test-presentation.pptx")


def setHealthStatus(uc):
    uc.setHealthCheckStatus("FSO_HAM_ONB_1", "pass (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_ONB_2", "pass (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_ONB_3", "fail (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_ONB_4", "fail (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_IMP_1", "pass (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_IMP_2", "pass (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_IMP_3", "pass (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_IMP_4", "pass (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_ADO_1", "Fail (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_ADO_2", "pass (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_OPT_1", "pass (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_OPT_2", "Fail (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_OPT_3", "pass (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_OPT_4", "Fail (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_USE_1", "Fail (xxx hc failed)")
    uc.setHealthCheckStatus("FSO_HAM_USE_2", "pass (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_ENG_1", "fail (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_ENG_2", "pass (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_ENG_3", "pass (xxx hc )")
    uc.setHealthCheckStatus("FSO_HAM_ENG_4", "fail (xxx hc )")


def testCxPptFsoUseCases_secondary_remediation_generation(uc, excels) :
    template = search("HybridApplicationMonitoringUseCase_template.pptx", "../")
    root = Presentation(template)
    data = uc.pitstop_data
    setHealthStatus(uc)
    fsoppt.markRaceTrackFailures(root,uc)

    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "onboard", fsoppt.SlideId.ONBOARD_REMEDIATION_SECONDARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "implement", fsoppt.SlideId.IMPLEMENT_REMEDIATION_SECONDARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "engage", fsoppt.SlideId.ENGAGE_REMEDIATION_SECONDARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "adopt", fsoppt.SlideId.ADOPT_REMEDIATION_SECONDARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "optimize", fsoppt.SlideId.OPTIMIZE_REMEDIATION_SECONDARY)
    fsoppt.generateRemediationSlides("folder_xxx", root, uc, excels, "use", fsoppt.SlideId.USE_REMEDIATION_SECONDARY)

    root.save(f"cx-ham-usecase-test-presentation-secondary.pptx")

def testGetPitstopTasks(uc):
    # onboard has 4 sequence or tasks
    assert len(uc.getPitstopTasks("onboard")) == 4
    # implement has 4 sequence or tasks
    assert len(uc.getPitstopTasks("implement")) == 4
    # use has 4 sequence or tasks
    assert len(uc.getPitstopTasks("use")) == 4
    # engage has 4 sequence or tasks
    assert len(uc.getPitstopTasks("engage")) == 4
    # adopt has 2 sequence or tasks
    assert len(uc.getPitstopTasks("adopt")) == 2
    # optimize has 4 sequence or tasks
    assert len(uc.getPitstopTasks("optimize")) == 4


    assert uc.getPitstopTasks("onboard")[0] == "FSO_HAM_ONB_1"
    assert uc.getPitstopTasks("onboard")[1] == "FSO_HAM_ONB_2"
    assert uc.getPitstopTasks("onboard")[2] == "FSO_HAM_ONB_3"
    assert uc.getPitstopTasks("onboard")[3] == "FSO_HAM_ONB_4"


    assert uc.getPitstopTasks("implement")[0] == "FSO_HAM_IMP_1"
    assert uc.getPitstopTasks("implement")[1] == "FSO_HAM_IMP_2"
    assert uc.getPitstopTasks("implement")[2] == "FSO_HAM_IMP_3"
    assert uc.getPitstopTasks("implement")[3] == "FSO_HAM_IMP_4"


    assert uc.getPitstopTasks("use")[0] == "FSO_HAM_USE_1"
    assert uc.getPitstopTasks("use")[1] == "FSO_HAM_USE_2"
    assert uc.getPitstopTasks("use")[2] == "FSO_HAM_USE_3"
    assert uc.getPitstopTasks("use")[3] == "FSO_HAM_USE_4"

    assert uc.getPitstopTasks("engage")[0] == "FSO_HAM_ENG_1"
    assert uc.getPitstopTasks("engage")[1] == "FSO_HAM_ENG_2"
    assert uc.getPitstopTasks("engage")[2] == "FSO_HAM_ENG_3"
    assert uc.getPitstopTasks("engage")[3] == "FSO_HAM_ENG_4"

    assert uc.getPitstopTasks("adopt")[0] == "FSO_HAM_ADO_1"
    assert uc.getPitstopTasks("adopt")[1] == "FSO_HAM_ADO_2"
    try:
        assert uc.getPitstopTasks("adopt")[2] == "FSO_HAM_ADO_3"
    except IndexError as e:
        assert e.args[0] == "list index out of range"

    assert uc.getPitstopTasks("optimize")[0] == "FSO_HAM_OPT_1"
    assert uc.getPitstopTasks("optimize")[1] == "FSO_HAM_OPT_2"
    assert uc.getPitstopTasks("optimize")[2] == "FSO_HAM_OPT_3"
    assert uc.getPitstopTasks("optimize")[3] == "FSO_HAM_OPT_4"

def testExecSheetsClass():
    folder = f"output/test_ham_report"
    ma_apm = 'test_ham_report-MaturityAssessment-apm.xlsx'
    ma_apm_raw = 'test_ham_report-MaturityAssessmentRaw-apm.xlsx'
    agent_matrix = 'test_ham_report-AgentMatrix.xlsx'
    xs = ExcelSheets(folder,(ma_apm, ma_apm_raw, agent_matrix))
    assert len(xs.getWorkBooks()) == 3
    assert len(xs.getHeaders(ma_apm, "AppAgentsAPM")) == 7
    print(xs.getHeaders(ma_apm, "AppAgentsAPM"))
    assert len(xs.getHeaders(agent_matrix, "Overall - Machine Agents")) == 4

    t1 = xs.getColumnTotal(ma_apm, "AppAgentsAPM", "percentAgentsReportingData")
    assert t1 == 400
    t1 = xs.getColumnAverage(ma_apm, "AppAgentsAPM", "percentAgentsReportingData")
    assert t1 == 11.11111111111111
    t1 = xs.getRowCountForColumnValue(
        ma_apm, "HealthRulesAndAlertingAPM", "numberOfActionsBoundToEnabledPolicies", "==", 1)
    assert t1 == 5

    expression = "BT lockdown total is <eval>count(btLockdownEnabled == True)</eval>"
    sub_exp, column_name, operator, value, is_eval_expr = xs.parseExpression(expression)
    if sub_exp is not None:
        count = xs.getValue(ma_apm, "BusinessTransactionsAPM", column_name, operator, value)
        assert count == 4
        assert xs.substituteExpression(expression, sub_exp, count) == "BT lockdown total is 4"
        print(xs.substituteExpression(expression, sub_exp, count))

    expression = "BT total is <eval>count(numberOfBTs == 3)</eval>"
    sub_exp, column_name, operator, value, is_eval_expr = xs.parseExpression(expression)
    count = xs.getValue(ma_apm, "BusinessTransactionsAPM", column_name, operator, value)
    assert count == 2
    assert xs.substituteExpression(expression, sub_exp, count) == "BT total is 2"

    expression = "BT total is <eval>count(numberOfBTs == 3)</eval>"
    sub_exp, column_name, operator, value, is_eval_expr = xs.parseExpression(expression)
    count = xs.getValue(ma_apm, "BusinessTransactionsAPM", column_name, operator, value)
    assert count == 2
    assert xs.substituteExpression(expression, sub_exp, count) == "BT total is 2"

    expression = "deployed agent count for 23+ is <eval>count(agentVersion == 23.3.0.0 compatible with 4.4.1.0)</eval>"
    sub_exp, column_name, operator, value, is_eval_expr = xs.parseExpression(expression)
    count = xs.getValue(agent_matrix, "Individual - appServerAgents" , column_name, operator, value)
    assert count == 1
    assert xs.substituteExpression(expression, sub_exp, count) == "deployed agent count for 23+ is 1"

    # Ensure non-eval expressions are not parsed
    expression = "Checkout <a href='https://community.cisco.com/t5/all-guides-for-hybrid-application-monitoring/hybrid-application-monitoring-guided-resources/ta-p/4679489'>Hybrid Application Monitoring Guided Resources</a> for a list of actions."
    sub_exp, column_name, operator, value, is_eval_expr = xs.parseExpression(expression)
    assert is_eval_expr == False

    # try == operator
    expression = "Ensure you keep your AppAgent versions up to date. Count of applications with all the AppAgents less than 1 years old:  <eval>count(percentAgentsLessThan1YearOld == 100)</eval> of your agents are less than 1 year old."
    sub_exp, column_name, operator, value, is_eval_expr = xs.parseExpression(expression)
    assert is_eval_expr == True
    count = xs.getValue(ma_apm, "AppAgentsAPM" , column_name, operator, value)
    assert count == 2
    assert xs.substituteExpression(expression, sub_exp, count) == "Ensure you keep your AppAgent versions up to date. Count of applications with all the AppAgents less than 1 years old:  2 of your agents are less than 1 year old."

    # try > operator
    expression = 'Most applications have more than one agent attached to their process/vms. Currently you have <eval>count(numberOfAgentsReportingData > 1)</eval> applications with more than one agent reporting data.'
    sub_exp, column_name, operator, value, is_eval_expr = xs.parseExpression(expression)
    assert is_eval_expr == True
    count = xs.getValue(ma_apm_raw, "AppAgentsAPM" , column_name, operator, value)
    assert count == 2
    assert xs.substituteExpression(expression, sub_exp, count) == "Most applications have more than one agent attached to their process/vms. Currently you have 2 applications with more than one agent reporting data."


    wb_name, sheet_name = xs.findSheetByHeader("percentAgentsReportingData")
    assert wb_name == ma_apm
    assert sheet_name == "AppAgentsAPM"

    wb_name, sheet_name = xs.findSheetByHeader("numberOfAgentsReportingData")
    print(wb_name, sheet_name)
