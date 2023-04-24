import json
import logging
import os
import re
from enum import Enum
from typing import Dict, Optional

from openpyxl import load_workbook
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt, Inches


class UseCase(tuple):
    def __init__(self, file: str):

        self.task_id_to_slide = {}
        self.task_id_to_health_check_status = {}

        ham_data_file = search(file, "../")

        with open(ham_data_file, 'r') as file:
            data = json.load(file)
            self.pitstop_data = data['pitstop']
        self._initSlideMapping()

    def pitstop_data(self) -> Dict:
        return self.pitstop_data

    def __str__(self) -> str:
        return self.pitstop_data.values()

    def _initSlideMapping(self):
        self.setSlideId('INTRO', 0)  # Welcome slide page 1
        self.setSlideId('RACETRACK', 1)  # Welcome slide page 2
        self.setSlideId("onboard", 2)
        self.setSlideId("onboard_remediation", 3)
        self.setSlideId("implement", 4)
        self.setSlideId("implement_remediation", 5)
        self.setSlideId("use", 6)
        self.setSlideId("use_remediation", 7)
        self.setSlideId("engage", 8)
        self.setSlideId("engage_remediation", 9)
        self.setSlideId("adopt", 10)
        self.setSlideId("adopt_remediation", 11)
        self.setSlideId("optimize", 12)
        self.setSlideId("optimize_remediation", 13)
        self.setSlideId("FINAL", 14)

    def getAllSlideIndexes(self):
        return list(self.task_id_to_slide.values())

    def getAllHealthCheckValues(self):
        return list(self.task_id_to_health_check_status.values())

    def isFailed(self, pitstop: str) -> bool:
        tasks = self.getPitstopTasks(pitstop)
        return any('fail' in self.task_id_to_health_check_status[task].lower() for task in tasks if task in self.task_id_to_health_check_status)

    def pitStopContainsFailureOrManualCheck(self, pitstop: str) -> bool:
        tasks = self.getPitstopTasks(pitstop)
        for task in tasks:
            if task in self.task_id_to_health_check_status \
                    and \
                    'fail' in self.task_id_to_health_check_status[task].lower() \
                    or \
                    'manual check' in self.task_id_to_health_check_status[task].lower():
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
                        'remediation_items': task_data.get('remediation_steps', None)
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

    # Set fixed row height
    for row_index, row in enumerate(table.rows):
        if row_index != 0:  # Skip the title row
            row.height = Inches(1)

    pass_mark = search("checkmark.png", "../")
    fail_mark = search("xmark.png", "../")

    for i, row in enumerate(data):
        for j, cell in enumerate(row):
            cell_obj = table.cell(i, j)
            cell_obj.text = str(cell)

            image_left = Inches(left+2) + table.columns[j].width  * j
            image_top = Inches(top) + table.rows[i].height * i
            # add marker image
            if "pass" in str(cell).lower() or "manual check" in str(cell).lower():
                slide.shapes.add_picture(pass_mark, image_left, image_top, width=Inches(0.2), height=Inches(0.2))
            if "fail" in str(cell).lower():
                slide.shapes.add_picture(fail_mark, image_left, image_top, width=Inches(0.2), height=Inches(0.2))

            for paragraph in cell_obj.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(fontSize)
                    run.font.color.rgb = color.value

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
    pPr.set('marL', '171450')
    pPr.set('indent', '171450')
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
        if idx % 3 == 0:
            run.text += part
        elif idx % 3 == 1:
            hyperlink_url = part
        else:
            run.text += part
            hlink = run.hyperlink
            hlink.address = hyperlink_url


def addRemediationTable(slide, data, color=None, fontSize=16, left=1.5, top=3.5, width=9.5, height=1.5):
    shape = slide.shapes.add_table(len(data), len(data[0]), Inches(left), Inches(top), Inches(width), Inches(height))
    table = shape.table
    table.columns[0].width = Inches(1.5)
    table.columns[2].width = Inches(2)
    table.columns[2].width = Inches(7)
    hyperlink_pattern = re.compile(r'<a href=[\'"](.*?)[\'"]>(.*?)<\/a>')

    for i, row in enumerate(data):
        for j, cell in enumerate(row):
            cell_text = str(cell)
            text_frame = table.cell(i, j).text_frame
            # Clear any existing paragraphs
            text_frame.clear()
            if "\n" not in cell_text:
                addCell(table.cell(i, j), cell_text, color, fontSize)
                continue

            for idx, line in enumerate(cell_text.split("\n")):
                segments = re.split(hyperlink_pattern, line)

                if idx == 0:
                    paragraph = text_frame.paragraphs[0]
                else:
                    paragraph = text_frame.add_paragraph()

                # paragraph = text_frame.add_paragraph()
                for i, segment in enumerate(segments):
                    if i % 3 == 0:
                        paragraph.add_run().text = segment
                    elif i % 3 == 1:  # url
                        url = segment
                    elif i % 3 == 2:  # hyperlink text
                        hyperlink_run = paragraph.add_run()
                        hyperlink_run.hyperlink.address = url
                        hyperlink_run.text = segment

                makeParaBulletPointed(paragraph)
                paragraph.font.size = Pt(8)


def getValuesInColumn(sheet, col1_value):
    values = []
    for column_cell in sheet.iter_cols(1, sheet.max_column):
        if column_cell[0].value == col1_value:
            j = 0
            for data in column_cell[1:]:
                values.append(data.value)
            break
    return values


def getValuesInColumnForController(sheet, col1_value, controller):
    values = []
    col1_index = -1
    col2_index = -1

    col2_value = "controller"

    # Find the indexes of col1_value and col2_value in the header row
    for i, cell in enumerate(sheet[1]):
        if cell.value == col1_value:
            col1_index = i
        if cell.value == col2_value:
            col2_index = i
        if col1_index != -1 and col2_index != -1:
            break

    # Check if both columns were found
    if col1_index == -1 or col2_index == -1:
        return values

    # Iterate through the rows in the worksheet, starting from row 2 to skip the header row
    for row in sheet.iter_rows(min_row=2):
        if row[col2_index].value == controller:
            values.append(row[col1_index].value)

    return values


def getRowCountForController(sheet, controller: str) -> int:
    CONTROLLER_NAME_COLUMN_INDEX = 1
    matching_rows = sum(row[0] == controller for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=CONTROLLER_NAME_COLUMN_INDEX, max_col=CONTROLLER_NAME_COLUMN_INDEX, values_only=True))
    return matching_rows


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

def add_image(slide, image_path, left, top, width, height):
    return slide.shapes.add_picture(image_path, left, top, width, height)


def search(filename, search_path="."):
    for root, _, files in os.walk(search_path):
        if filename in files:
            return os.path.join(root, filename)
    return None

def markRaceTrackFailures(root, uc: UseCase):
    # racetrack slide is the second slide
    slide = root.slides[1]
    pass_mark = search("checkmark.png", "../")
    fail_mark = search("xmark.png", "../")
    image_width = Inches(0.3)
    image_height = Inches(0.3)

    # IMPORTANT - EXTERNAL DEPENDENCY each pitstop has a named shape in the
    # PPTX template and is used for the iteration below else pitstop shapes won't be found
    # and will not be visually marked for fail/pass checkmarks
    pitstops = ("onboard","use","implement","optimize","adopt","engage")

    shapes_dict = {shape.name: shape for shape in slide.shapes}
    for pitstop in pitstops:
        try:
            if uc.isFailed(pitstop):
                _ = add_image(slide, fail_mark , shapes_dict[pitstop].left, shapes_dict[pitstop].top, image_width, image_height)
            else:
                _ = add_image(slide, pass_mark , shapes_dict[pitstop].left, shapes_dict[pitstop].top, image_width, image_height)
        except KeyError:
            logging.error(f"Shape with name '{pitstop}' not found in the slide. This prevents properly marking the race track with visual markers. ")

def createCxHamUseCasePpt(folder: str):
    logging.info(f"Creating CX HAM Use Case PPT for {folder}")
    apm_wb = load_workbook(f"output/{folder}/{folder}-MaturityAssessment-apm.xlsx")
    db_wb = load_workbook(f"output/{folder}/{folder}-AgentMatrix.xlsx")
    uc = UseCase('backend/resources/pptAssets/HybridApplicationMonitoringUseCase.json')
    _ = calculate_kpis(apm_wb, db_wb, uc)
    root = Presentation("backend/resources/pptAssets/HybridApplicationMonitoringUseCase_template.pptx")

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
    markRaceTrackFailures(root,uc)


    root.save(f"output/{folder}/{folder}-cx-HybridApplicationMonitoringUseCaseMaturityAssessment-presentation.pptx")


def generatePitstopHealthCheckTable(folder, root, uc, pitstop):
    slide = root.slides[uc.getSlideId(pitstop)]
    data = [["Controller", "Checklist Item", "Tooltips", "Exit Criteria Logic"]]
    for task in uc.getPitstopTasks(pitstop):
        data.append([folder, f"{uc.getChecklistItem(task)}", f"{uc.getToolTip(task)}", uc.getHealthCheckStatus(task)])
    addTable(slide, data, fontSize=10, top=2, left=1.5)


def generateRemediationSlides(folder: str, root: Presentation, uc: UseCase, pitstop: str, remediation_slide: str):
    if uc.pitStopContainsFailureOrManualCheck(pitstop):
        slide = root.slides[uc.getSlideId(remediation_slide)]
        data = [["Controller", "Checklist Item", "Recommendation"]]
        for task in uc.getPitstopTasks(pitstop):
            data.append(
                [folder,
                 f"{uc.getChecklistItem(task)}",
                 '\n'.join(value["remediation_item"] for value in uc.getRemediationList(task).values())
                 ]
            )
        addRemediationTable(slide, data, fontSize=10, top=2, left=.5)


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
    if not uc.pitStopContainsFailureOrManualCheck("onboard"):
        slides_to_keep.remove(uc.getSlideId("onboard_remediation"))
    if not uc.pitStopContainsFailureOrManualCheck("implement"):
        slides_to_keep.remove(uc.getSlideId("implement_remediation"))
    if not uc.pitStopContainsFailureOrManualCheck("use"):
        slides_to_keep.remove(uc.getSlideId("use_remediation"))
    if not uc.pitStopContainsFailureOrManualCheck("engage"):
        slides_to_keep.remove(uc.getSlideId("engage_remediation"))
    if not uc.pitStopContainsFailureOrManualCheck("adopt"):
        slides_to_keep.remove(uc.getSlideId("adopt_remediation"))
    if not uc.pitStopContainsFailureOrManualCheck("optimize"):
        slides_to_keep.remove(uc.getSlideId("optimize_remediation"))

    filter_slides(slides_to_keep, root)


def calculate_kpis(apm_wb, agent_wb, uc: UseCase):
    # currently only supports one controller report out of the workbook
    controller = getValuesInColumn(apm_wb["Analysis"], "controller")[0]
    logging.info(f"processing report for 1st controller only as multiple controllers are not supported yet: {controller}")

    totalApplications = getRowCountForController(apm_wb["Analysis"], controller)
    percentAgentsReportingData = getValuesInColumnForController(apm_wb["AppAgentsAPM"], "percentAgentsReportingData", controller)
    countOfpercentAgentsReportingData = len([x for x in percentAgentsReportingData if x >= 1])
    numberOfBTsList = getValuesInColumnForController(apm_wb["BusinessTransactionsAPM"], "numberOfBTs", controller)
    customMatchRulesList = getValuesInColumnForController(apm_wb["BusinessTransactionsAPM"], "numberCustomMatchRules", controller)
    numberOfDataCollectorsConfigured = getValuesInColumnForController(apm_wb["DataCollectorsAPM"], "numberOfDataCollectorFieldsConfigured", controller)
    countOfNumberOfDataCollectorsConfigured = len([x for x in numberOfDataCollectorsConfigured if x >= 1])
    numberOfCustomHealthRules = getValuesInColumnForController(apm_wb["HealthRulesAndAlertingAPM"], "numberOfCustomHealthRules", controller)
    numberOfHealthRuleViolations = getValuesInColumnForController(apm_wb["HealthRulesAndAlertingAPM"], "numberOfHealthRuleViolations", controller)
    numberOfDefaultHealthRulesModified = getValuesInColumnForController(apm_wb["HealthRulesAndAlertingAPM"], "numberOfDefaultHealthRulesModified", controller)
    numberOfActionsBoundToEnabledPoliciesList = getValuesInColumnForController(apm_wb["HealthRulesAndAlertingAPM"], "numberOfActionsBoundToEnabledPolicies", controller)
    dashboardsList = getValuesInColumnForController(apm_wb["DashboardsAPM"], "numberOfDashboards", controller)

    dbAgentList = getValuesInColumnForController(agent_wb["Individual - dbAgents"], "status", controller)
    dbAgentsActiveCount = len([x for x in dbAgentList if x == "ACTIVE"])

    ### ONB
    FSO_HAM_ONB_1 = f'manual check'
    FSO_HAM_ONB_2 = f'manual check'
    FSO_HAM_ONB_3 = f'manual check'
    FSO_HAM_ONB_4 = f'manual check'

    ### IMP
    FSO_HAM_IMP_1 = f'manual check'
    FSO_HAM_IMP_2 = f'Pass ({totalApplications} applications on-boarded)' if totalApplications >= 1 else f'Fail. At least one application needs to be instrumented and reporting metric data into AppDynamics controller'
    FSO_HAM_IMP_3 = f'manual check'

    # FSO_HAM_IMP_4 = f'Pass ({customMatchRulesList})' if all(countOfCustomRule >= 2 for countOfCustomRule in customMatchRulesList) else f'Fail (Not all applications have at least 2 custom BT match rules). Only {len([x for x in customMatchRulesList if x >= 2])} have at least 2 custom match rules out of {totalApplications}.'
    FSO_HAM_IMP_4 = f'Pass' if countOfpercentAgentsReportingData == totalApplications else f'Fail (Not all application agents are reporting data). Only  {countOfpercentAgentsReportingData} reporting data out of {totalApplications}.'

    ### USE
    # FSO_HAM_USE_1 = f'Pass' if countOfpercentAgentsReportingData == totalApplications else f'Fail (Not all application agents are reporting data). Only  {countOfpercentAgentsReportingData} reporting data out of {totalApplications}.'
    FSO_HAM_USE_1 = f'Pass ({customMatchRulesList})' if all(countOfCustomRule >= 2 for countOfCustomRule in customMatchRulesList) else f'Fail (Not all applications have at least 2 custom BT match rules). Only {len([x for x in customMatchRulesList if x >= 2])} have at least 2 custom match rules out of {totalApplications}.'

    FSO_HAM_USE_2 = f'Pass' if all(count >= 5 for count in numberOfBTsList) else f'Fail (Not all applications have at least 5 Business Transactions). Only {len([x for x in numberOfBTsList if x >= 5])} have at least 5 Business Transactions out of {totalApplications}.'

    FSO_HAM_USE_3 = f'Pass' if all(count >= 2 for count in numberOfCustomHealthRules) else f'Fail (At least 2 custom health rules must be configured per application). Only {len([count for count in numberOfCustomHealthRules if count >= 2])} applications have at least 2 custom health rules configured out of {totalApplications}.'
    FSO_HAM_USE_4 = f'Pass ({countOfNumberOfDataCollectorsConfigured} data collectors found)' if countOfNumberOfDataCollectorsConfigured >= 2 else 'Fail'

    ### ENG
    FSO_HAM_ENG_1 = f'Pass' if all(count >= 5 for count in numberOfActionsBoundToEnabledPoliciesList) else f'Fail (Not all applications have at least 2 policies with actionable alerts). Only {len([x for x in numberOfActionsBoundToEnabledPoliciesList if x >= 2])} applications have at least 2 actionable alerts out of {totalApplications} applications.'

    FSO_HAM_ENG_2 = f'Pass' if all(count >= 2 for count in dashboardsList) else f'Fail (there should be at least two dashboards configured). Only {len([x for x in dashboardsList if x >= 2])} dashboards are configured'
    FSO_HAM_ENG_3 = f'manual check'
    FSO_HAM_ENG_4 = f'manual check'

    ### ADO
    FSO_HAM_ADO_1 = f'manual check'
    FSO_HAM_ADO_2 = f'Pass' if dbAgentsActiveCount > 0 else f'Fail (there should be at least one Active Database Agent configured). Currently {dbAgentsActiveCount} Active Database agents are configured.'

    ### OPT
    FSO_HAM_OPT_1 = f'manual check'
    FSO_HAM_OPT_2 = f'manual check'
    FSO_HAM_OPT_3 = f'Pass' if all(count >= 6 for count in numberOfCustomHealthRules) else f'Fail (At least 6 custom health rules must be configured per applications). Only {len([count for count in numberOfCustomHealthRules if count >= 6])} application has 6 custom health rules configured.'
    FSO_HAM_OPT_4 = f'Pass' if all(count >= 10 for count in numberOfBTsList) else f'Fail (there should be at least 10 Business Translations detected per application). Only {len([count for count in numberOfBTsList if count >= 10])} applications have at least 10 Business Transaction'

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

    return kpi_dictionary
