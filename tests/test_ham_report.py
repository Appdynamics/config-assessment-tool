import pytest
from pptx import Presentation
from output.presentations.cxPptFsoUseCases import UseCase, generateRemediationSlides, cleanup_slides, generatePitstopHealthCheckTable, createCxHamUseCasePpt, markRaceTrackFailures
from output.presentations.cxPptFsoUseCases import UseCase


@pytest.fixture
def uc():
    return UseCase('../backend/resources/pptAssets/HybridApplicationMonitoringUseCase.json')

def testCxPptFsoUseCases(uc):
    root = Presentation("../backend/resources/pptAssets/HybridApplicationMonitoringUseCase_template.pptx")
    data = uc.pitstop_data
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

    markRaceTrackFailures(root,uc)

    generatePitstopHealthCheckTable("folder_xxx", root, uc, "onboard")
    generatePitstopHealthCheckTable("folder_xxx", root, uc, "implement")
    generatePitstopHealthCheckTable("folder_xxx", root, uc, "use")
    generatePitstopHealthCheckTable("folder_xxx", root, uc, "engage")
    generatePitstopHealthCheckTable("folder_xxx", root, uc, "adopt")
    generatePitstopHealthCheckTable("folder_xxx", root, uc, "optimize")

    generateRemediationSlides("folder_xxx", root, uc, "onboard", "onboard_remediation")
    generateRemediationSlides("folder_xxx", root, uc, "implement", "implement_remediation")
    generateRemediationSlides("folder_xxx", root, uc, "engage", "engage_remediation")
    generateRemediationSlides("folder_xxx", root, uc, "adopt", "adopt_remediation")
    generateRemediationSlides("folder_xxx", root, uc, "optimize", "optimize_remediation")
    generateRemediationSlides("folder_xxx", root, uc, "use", "use_remediation")
    cleanup_slides(root, uc)
    root.save(f"cx-ham-usecase-test-presentation.pptx")

def test_getPitstopTasks(uc):
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

@pytest.mark.skip(reason="need to refactor to work")
def test_ReportUsingDemoControllerOuput():
    createCxHamUseCasePpt("appd-cse")