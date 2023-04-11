import logging
import os
from pptx.oxml.xmlchemy import OxmlElement
import re
from enum import Enum

from openpyxl import load_workbook
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.slide import Slide
from pptx.util import Inches, Pt
from typing import List, Dict, Optional


import json
from typing import List, Dict, Optional

class UseCase(tuple):
    def __init__(self, json_file: str):

        self.task_id_to_slide = {}
        self.task_id_to_health_check_status = {}

        with open(json_file, 'r') as file:
            data = json.load(file)
            self.pitstop_data = data['pitstop']
        self._initSlideMapping()

    def pitstop_data(self) -> Dict:
        return self.pitstop_data
    def __str__(self) -> str:
        return self.pitstop_data.values()

    def _initSlideMapping(self):
        self.setSlideId('INTRO',0) # Welcome slide page 1
        self.setSlideId('RACETRACK',1) # Welcome slide page 2
        self.setSlideId("onboard",2)
        self.setSlideId("onboard_remediation",3)
        self.setSlideId("implement",4)
        self.setSlideId("implement_remediation",5)
        self.setSlideId("use",6)
        self.setSlideId("use_remediation",7)
        self.setSlideId("engage",8)
        self.setSlideId("engage_remediation",9)
        self.setSlideId("adopt",10)
        self.setSlideId("adopt_remediation",11)
        self.setSlideId("optimize",12)
        self.setSlideId("optimize_remediation",13)
        self.setSlideId("FINAL",14)
        print("initialized the slide deck index")
        print(self.task_id_to_slide)

    def getAllSlideIndexes(self):
        return list(self.task_id_to_slide.values())

    def getAllHealthCheckValues(self):
        return list(self.task_id_to_health_check_status.values())

    def pitStopContainsFailure(self, pitstop: str) -> bool:
        tasks = self.getPitstopTasks(pitstop)
        for task in tasks:
            if task in self.task_id_to_health_check_status \
                    and 'fail' in self.task_id_to_health_check_status[task].lower():
                return True

        return False

    def _get_task_data(self, task_id: str) -> Dict[str, Optional[str]]:
        for pitstop_stage in self.pitstop_data:
            for sequence_number, task_data in self.pitstop_data[pitstop_stage]['checklist_sequence'].items():
                if task_data['task_id'] == task_id:
                    return {
                        'checklist_item': task_data['checklist_item'],
                        'tooltip': task_data.get('tooltip', None),
                        'exit_criteria_logic': task_data.get('exit_criteria_logic', None),
                        'min_pass_threshold': task_data.get('min_pass_threshold', None),
                        'remediation_items' : task_data.get('remediation_steps', None)
                    }
        return {}


    def getAllTaskIds(self):
        task_ids = []
        for pitstop_stage in self.pitstop_data:
            for sequence_number, task_data in self.pitstop_data[pitstop_stage]['checklist_sequence'].items():
                task_ids.append(task_data['task_id'])
        return task_ids

    def getPitstopTasks(self, pitstop: str) -> Dict[str, str]:
        task_ids = []
        for pitstop_stage in self.pitstop_data[pitstop].values():
            for k in pitstop_stage.values():
                task_ids.append(k['task_id'])
        return task_ids if task_ids else None

    def getChecklistItem(self, task_id: str) -> Optional[str]:
        task_data = self._get_task_data(task_id)
        return task_data['checklist_item'] if task_data else None
    def getToolTip(self, task_id: str) -> Optional[str]:
        task_data = self._get_task_data(task_id)
        return task_data.get('tooltip', None) if task_data else None

    def getExitCriteriaLogic(self, task_id: str) -> Optional[str]:
        task_data = self._get_task_data(task_id)
        return task_data.get('exit_criteria_logic', None) if task_data else None

    def getMinPassThreshold(self, task_id: str) -> Optional[str]:
        task_data = self._get_task_data(task_id)
        return task_data.get('min_pass_threshold', None) if task_data else None

    def getRemediationList(self, task_id: str) -> Optional[list]:
        task_data = self._get_task_data(task_id)
        return task_data.get('remediation_items', None) if task_data else None

    def setSlideId(self, task_id: str, slide_index: int):
        self.task_id_to_slide[task_id] = slide_index

    def getSlideId(self, task_id: str) -> Optional[int]:
        # this avoids a potential KeyError
        return self.task_id_to_slide.get(task_id, None)
    def setHealthCheckStatus(self, task_id: str, hc: str):
        self.task_id_to_health_check_status[task_id] = hc

    def getHealthCheckStatus(self, task_id: str) -> Optional[str]:
        # this avoids a potential KeyError
        return self.task_id_to_health_check_status.get(task_id, None)

class Color(Enum):
    WHITE = RGBColor(255, 255, 255)
    BLACK = RGBColor(0, 0, 0)
    RED = RGBColor(255, 0, 0)
    GREEN = RGBColor(0, 255, 0)
    BLUE = RGBColor(0, 0, 255)

def addTable(slide, data, color: Color = Color.BLACK, fontSize: int = 16, left: int = 0.25, top: int = 3.5, width: int = 9.5, height: int = 1.5):
    shape = slide.shapes.add_table(len(data), len(data[0]), Inches(left), Inches(top), Inches(width), Inches(height))
    table = shape.table

    for i, row in enumerate(data):
        for j, cell in enumerate(row):
            table.cell(i, j).text = str(cell)
            for paragraph in table.cell(i, j).text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(fontSize)
                    run.font.color.rgb = color.value

import re
from pptx.util import Pt, Inches
from pptx.action import Hyperlink
from pptx.dml.color import ColorFormat
# from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.dml import MSO_THEME_COLOR


def addCell(cell, text, color=None, fontSize=16):
    cell.text = text
    for paragraph in cell.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(fontSize)
            if color:
                run.font.color.rgb = color.value

def SubElement(parent, tagname, **kwargs):
    element = OxmlElement(tagname)
    element.attrib.update(kwargs)
    parent.append(element)
    return element

def makeParaBulletPointed(para):
    """Bullets are set to Arial,
        actual text can be a different font"""
    pPr = para._p.get_or_add_pPr()
    ## Set marL and indent attributes
    pPr.set('marL','171450')
    pPr.set('indent','171450')
    ## Add buFont
    _ = SubElement(parent=pPr,
                   tagname="a:buFont",
                   typeface="Arial",
                   panose="020B0604020202020204",
                   pitchFamily="34",
                   charset="0"
                   )
    ## Add buChar
    _ = SubElement(parent=pPr,
                   tagname='a:buChar',
                   char="•")

def addHyperLinkCell(run, cell_text):
    hyperlink_pattern = re.compile(r'<a href=[\'"](.*?)[\'"]>(.*?)<\/a>')
    parts = hyperlink_pattern.split(cell_text)
    for idx, part in enumerate(parts):
        print(idx, part)
        if idx % 3 == 0:
            run.text += part
        elif idx % 3 == 1:
            hyperlink_url = part
        else:
            run.text += part
            hlink = run.hyperlink
            hlink.address = hyperlink_url
            run.font.color.theme_color = MSO_THEME_COLOR.HYPERLINK

def addRemediationTable(slide, data, color=None, fontSize=16, left=1.5, top=3.5, width=9.5, height=1.5):
    shape = slide.shapes.add_table(len(data), len(data[0]), Inches(left), Inches(top), Inches(width), Inches(height))
    table = shape.table
    table.columns[0].width = Inches(1.5)
    table.columns[2].width = Inches(2)
    table.columns[2].width = Inches(7)
    hyperlink_pattern = re.compile(r'<a href=[\'"](.*?)[\'"]>(.*?)<\/a>')

    for i, row in enumerate(data):
        for j, cell in enumerate(row):
            print("row {}, col {}, cell: {}".format(i, j, cell))
            cell_text = str(cell)
            text_frame = table.cell(i, j).text_frame
            # Clear any existing paragraphs
            text_frame.clear()
            if "\n" not in cell_text:
                addCell(table.cell(i, j), cell_text, color, fontSize)
                continue

            for line in cell_text.split("\n"):
                if hyperlink_pattern.search(line):
                    print("Matched hyperlink pattern")
                    p = text_frame.add_paragraph()
                    run = p.add_run()
                    addHyperLinkCell(run, line)
                    makeParaBulletPointed(p)
                    p.font.size = Pt(12)
                else:
                    print(f"Adding regular text line: {line}")
                    p = text_frame.add_paragraph()
                    p.text = line
                    makeParaBulletPointed(p)
                    p.font.size = Pt(12)


def getValuesInColumn(sheet, param):
    values = []
    for column_cell in sheet.iter_cols(1, sheet.max_column):
        if column_cell[0].value == param:
            j = 0
            for data in column_cell[1:]:
                values.append(data.value)
            break
    return values


def getAppsWithScore(sheet, assessmentScore):
    values = []
    for column_cell in sheet.iter_cols(1, sheet.max_column):
        if column_cell[0].value == "OverallAssessment":
            j = 0
            for idx, data in enumerate(column_cell[1:]):
                if data.value == assessmentScore:
                    values.append(sheet[f"C{idx + 2}"].value)  # +2 because of the header
            break
    return values


def createCxHamUseCasePpt(folder: str):
    logging.info(f"Creating CX HAM Use Case PPT for {folder}")
    wb = load_workbook(f"output/{folder}/{folder}-MaturityAssessment-apm.xlsx")
    uc = UseCase('backend/resources/pptAssets/HybridApplicationMonitoringUseCase.json')
    metrics = calculate_kpis(wb, uc);
    root = Presentation("backend/resources/pptAssets/HybridApplicationMonitoringUseCase_template.pptx")

    # Example usage
    print("------------------------------------")
    print(uc.getChecklistItem('FSO_HAM_ONB_1'))
    print(uc.getToolTip('FSO_HAM_ONB_1'))
    print(uc.getExitCriteriaLogic('FSO_HAM_ONB_1'))
    print("------------------------------------")
    print("------------------------------------")
    print(uc.getChecklistItem('FSO_HAM_ONB_2'))
    print(uc.getToolTip('FSO_HAM_ONB_2'))
    print(uc.getExitCriteriaLogic('FSO_HAM_ONB_2'))
    print("------------------------------------")

    ############################# Onboard ###########################
    generatePitstopHealthCheckTable(folder, root, uc, "onboard")
    generateRemediationSlides(folder, root, uc, "onboard", "onboard_remediation")
    ############################# Implement ###########################
    generatePitstopHealthCheckTable(folder, root, uc, "implement")
    generateRemediationSlides(folder, root, uc, "implement", "implement_remediation")
    ############################ Use ##################################
    generatePitstopHealthCheckTable(folder, root, uc, "use")
    generateRemediationSlides(folder, root, uc, "use", "use_remediation")
    ############################ Engage ###############################
    generatePitstopHealthCheckTable(folder, root, uc, "engage")
    generateRemediationSlides(folder, root, uc, "engage", "engage_remediation")
    ############################ Adopt ###############################
    generatePitstopHealthCheckTable(folder, root, uc, "adopt")
    generateRemediationSlides(folder, root, uc, "adopt", "adopt_remediation")
    ########################### Optimize ##############################
    generatePitstopHealthCheckTable(folder, root, uc, "optimize")
    generateRemediationSlides(folder, root, uc, "optimize", "optimize_remediation")



    cleanup_slides(root, uc)
    root.save(f"output/{folder}/{folder}-cx-HybridApplicationMonitoringUseCaseMaturityAssessment-presentation.pptx")


def generatePitstopHealthCheckTable(folder, root, uc, pitstop):
    slide = root.slides[uc.getSlideId(pitstop)]
    data = [ ["Controller", "Checklist Item", "Tooltips", "Exit Criteria Logic"] ]
    for task in uc.getPitstopTasks(pitstop):
        data.append([folder, f"{uc.getChecklistItem(task)}", f"{uc.getToolTip(task)}", uc.getHealthCheckStatus(task)])
    addTable(slide, data, fontSize=10, top=2, left=1.5)


def generateRemediationSlides(folder: str, root: Presentation, uc : UseCase, pitstop: str, remediation_slide: str):
    if uc.pitStopContainsFailure(pitstop):
        slide = root.slides[uc.getSlideId(remediation_slide)]
        data = [ ["Controller", "Checklist Item", "Recommendation"] ]
        for task in uc.getPitstopTasks(pitstop):
            data.append(
                [folder,
                 f"{uc.getChecklistItem(task)}",
                 '\n'.join( value["remediation_item"] for value in uc.getRemediationList(task).values())
                 ]
            )
        addRemediationTable(slide , data, fontSize=10, top=2, left=.5)

def filter_slides(keep_slide_indexes, prs: Presentation):
    total_slides = len(prs.slides)
    slides_to_remove = {i for i in range(total_slides) if i not in keep_slide_indexes}
    for slide_index in sorted(slides_to_remove, reverse=True):
        slide = prs.slides[slide_index]
        id_to_index_mapping = {slide.id: (i, slide.rId) for i, slide in enumerate(prs.slides._sldIdLst)}
        slide_id = slide.slide_id
        slide_rId = id_to_index_mapping[slide_id][1]
        prs.part.drop_rel(slide_rId)
        del prs.slides._sldIdLst[id_to_index_mapping[slide_id][0]]
    return prs


def cleanup_slides(root: Presentation, uc: UseCase):
    slides_to_keep = uc.getAllSlideIndexes()
    if not uc.pitStopContainsFailure("onboard"):
        slides_to_keep.remove(uc.getSlideId("onboard_remediation"))
    if not uc.pitStopContainsFailure("implement"):
        slides_to_keep.remove(uc.getSlideId("implement_remediation"))
    if not uc.pitStopContainsFailure("use"):
        slides_to_keep.remove(uc.getSlideId("use_remediation"))
    if not uc.pitStopContainsFailure("engage"):
        slides_to_keep.remove(uc.getSlideId("engage_remediation"))
    if not uc.pitStopContainsFailure("adopt"):
        slides_to_keep.remove(uc.getSlideId("adopt_remediation"))
    if not uc.pitStopContainsFailure("optimize"):
        slides_to_keep.remove(uc.getSlideId("optimize_remediation"))

    print("Keeping slides: ", slides_to_keep)
    filter_slides(slides_to_keep, root)

def calculate_kpis(wb: str, uc: UseCase):

    totalApplications = wb["Analysis"].max_row - 1
    percentAgentsReportingData = getValuesInColumn(wb["AppAgentsAPM"], "percentAgentsReportingData")
    countOfpercentAgentsReportingData = len([x for x in percentAgentsReportingData if x >= 1])
    numberOfBTs= getValuesInColumn(wb["BusinessTransactionsAPM"], "numberOfBTs")
    countOfNumberOfBTs = len([x for x in numberOfBTs if x >= 1])
    numberOfCustomHealthRules= getValuesInColumn(wb["HealthRulesAndAlertingAPM"], "numberOfCustomHealthRules")
    countOfNumberOfCustomHealthRules = len([x for x in numberOfCustomHealthRules if x >= 1])
    numberOfDataCollectorsConfigured= getValuesInColumn(wb["DataCollectorsAPM"], "numberOfDataCollectorFieldsConfigured")
    countOfNumberOfDataCollectorsConfigured = len([x for x in numberOfDataCollectorsConfigured if x >= 1])

    percentAgentsReportingData = getValuesInColumn(wb["AppAgentsAPM"], "percentAgentsReportingData")
    countOfpercentAgentsReportingData = len([x for x in percentAgentsReportingData if x >= 1])
    numberOfBTs= getValuesInColumn(wb["BusinessTransactionsAPM"], "numberOfBTs")
    countOfNumberOfBTs = len([x for x in numberOfBTs if x >= 1])
    numberOfCustomHealthRules= getValuesInColumn(wb["HealthRulesAndAlertingAPM"], "numberOfCustomHealthRules")
    countOfNumberOfCustomHealthRules = len([x for x in numberOfCustomHealthRules if x >= 1])
    numberOfDataCollectorsConfigured= getValuesInColumn(wb["DataCollectorsAPM"], "numberOfDataCollectorFieldsConfigured")
    countOfNumberOfDataCollectorsConfigured = len([x for x in numberOfDataCollectorsConfigured if x >= 1])

    numberOfHealthRuleViolations = getValuesInColumn(wb["HealthRulesAndAlertingAPM"], "numberOfHealthRuleViolations")
    numberOfDefaultHealthRulesModified = getValuesInColumn(wb["HealthRulesAndAlertingAPM"], "numberOfDefaultHealthRulesModified")
    numberOfActionsBoundToEnabledPolicies = getValuesInColumn(wb["HealthRulesAndAlertingAPM"], "numberOfActionsBoundToEnabledPolicies")
    countOfActionsBoundToEnabledPolicies = len([x for x in numberOfActionsBoundToEnabledPolicies if x >= 1])
    dashboardsList = getValuesInColumn(wb["DashboardsAPM"], "numberOfDashboards")
    countOfDashboardsList = len([x for x in dashboardsList if x >= 1])

    numberOfCustomHealthRules = getValuesInColumn(wb["HealthRulesAndAlertingAPM"], "numberOfCustomHealthRules")


    FSO_HAM_ONB_1 = f'Manual'
    FSO_HAM_ONB_2 = f'Manual'
    FSO_HAM_ONB_3 = f'Manual'
    FSO_HAM_ONB_4 = f'Manual'
    FSO_HAM_USE_1 = f'Pass ({countOfpercentAgentsReportingData})' if countOfpercentAgentsReportingData>= 1 else 'Fail'
    FSO_HAM_USE_2 = f'Pass ({countOfNumberOfBTs})' if countOfNumberOfBTs >=1 else 'Fail'
    FSO_HAM_USE_3 = f'Pass ({countOfNumberOfCustomHealthRules})' if countOfNumberOfCustomHealthRules>=5 else 'Fail'
    FSO_HAM_USE_4 = f'Pass ({countOfNumberOfCustomHealthRules})' if countOfNumberOfCustomHealthRules>=5 else 'Fail'
    FSO_HAM_IMP_1 = f'TBI'
    FSO_HAM_IMP_2 = f'Pass ({totalApplications})' if totalApplications >= 1 else 'Fail'
    FSO_HAM_IMP_3 = f'TBI'
    FSO_HAM_IMP_4 = f'TBI'
    FSO_HAM_ENG_1 = f'Pass({countOfActionsBoundToEnabledPolicies})' if totalApplications == countOfActionsBoundToEnabledPolicies else 'Fail'
    FSO_HAM_ENG_2 = f'Pass({countOfDashboardsList})' if countOfDashboardsList >=2 else 'Fail'
    FSO_HAM_ENG_3 = f'Pass({countOfNumberOfCustomHealthRules})' if countOfNumberOfCustomHealthRules>=5 else 'Fail'
    FSO_HAM_ENG_4 = f'TBI'
    FSO_HAM_ADO_1 = f'TBI'
    FSO_HAM_ADO_2 = f'FAIL xxxx'
    FSO_HAM_OPT_1 = f'TBI'
    FSO_HAM_OPT_2 = f'FAIL yyyy '
    FSO_HAM_OPT_3 = f'TBI'
    FSO_HAM_OPT_4 = f'FAIL zzzz'

    uc.setHealthCheckStatus('FSO_HAM_ONB_1', FSO_HAM_ONB_1)
    uc.setHealthCheckStatus('FSO_HAM_ONB_2', FSO_HAM_ONB_2)
    uc.setHealthCheckStatus('FSO_HAM_ONB_3', FSO_HAM_ONB_3)
    uc.setHealthCheckStatus('FSO_HAM_ONB_4', FSO_HAM_ONB_4)
    uc.setHealthCheckStatus('FSO_HAM_USE_1', FSO_HAM_USE_1)
    uc.setHealthCheckStatus('FSO_HAM_USE_2', FSO_HAM_USE_2)
    uc.setHealthCheckStatus('FSO_HAM_USE_3', FSO_HAM_USE_3)
    uc.setHealthCheckStatus('FSO_HAM_USE_4', FSO_HAM_USE_4)
    uc.setHealthCheckStatus('FSO_HAM_IMP_1', FSO_HAM_IMP_1)
    uc.setHealthCheckStatus('FSO_HAM_IMP_2', FSO_HAM_IMP_2)
    uc.setHealthCheckStatus('FSO_HAM_IMP_3', FSO_HAM_IMP_3)
    uc.setHealthCheckStatus('FSO_HAM_IMP_4', FSO_HAM_IMP_4)
    uc.setHealthCheckStatus('FSO_HAM_ENG_1', FSO_HAM_ENG_1)
    uc.setHealthCheckStatus('FSO_HAM_ENG_2', FSO_HAM_ENG_2)
    uc.setHealthCheckStatus('FSO_HAM_ENG_3', FSO_HAM_ENG_3)
    uc.setHealthCheckStatus('FSO_HAM_ENG_4', FSO_HAM_ENG_4)
    uc.setHealthCheckStatus('FSO_HAM_ADO_1', FSO_HAM_ADO_1)
    uc.setHealthCheckStatus('FSO_HAM_ADO_2', FSO_HAM_ADO_2)
    uc.setHealthCheckStatus('FSO_HAM_OPT_1', FSO_HAM_OPT_1)
    uc.setHealthCheckStatus('FSO_HAM_OPT_2', FSO_HAM_OPT_2)
    uc.setHealthCheckStatus('FSO_HAM_OPT_3', FSO_HAM_OPT_3)
    uc.setHealthCheckStatus('FSO_HAM_OPT_4', FSO_HAM_OPT_4)




    kpi_dictionary = {
        'FSO_HAM_ONB_1': FSO_HAM_ONB_1,
        'FSO_HAM_ONB_2': FSO_HAM_ONB_2,
        'FSO_HAM_ONB_3': FSO_HAM_ONB_3,
        'FSO_HAM_ONB_4': FSO_HAM_ONB_4,
        'FSO_HAM_USE_1': FSO_HAM_USE_1,
        'FSO_HAM_USE_2': FSO_HAM_USE_2,
        'FSO_HAM_USE_3': FSO_HAM_USE_3,
        'FSO_HAM_USE_4': FSO_HAM_USE_4,
        'FSO_HAM_IMP_1': FSO_HAM_IMP_1,
        'FSO_HAM_IMP_2': FSO_HAM_IMP_2,
        'FSO_HAM_IMP_3': FSO_HAM_IMP_3,
        'FSO_HAM_IMP_4': FSO_HAM_IMP_4,
        'FSO_HAM_ENG_1': FSO_HAM_ENG_1,
        'FSO_HAM_ENG_2': FSO_HAM_ENG_2,
        'FSO_HAM_ENG_3': FSO_HAM_ENG_3,
        'FSO_HAM_ENG_4': FSO_HAM_ENG_4,
        'FSO_HAM_ADO_1': FSO_HAM_ADO_1,
        'FSO_HAM_ADO_2': FSO_HAM_ADO_2,
        'FSO_HAM_OPT_1': FSO_HAM_OPT_1,
        'FSO_HAM_OPT_2': FSO_HAM_OPT_2,
        'FSO_HAM_OPT_3': FSO_HAM_OPT_3,
        'FSO_HAM_OPT_4': FSO_HAM_OPT_4
    }



    print(kpi_dictionary)

    return kpi_dictionary



