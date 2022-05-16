import json
import logging
import os
from datetime import datetime
from enum import Enum
from typing import List

import click
from openpyxl import load_workbook
from tzlocal import get_localzone
from pptx import Presentation
from pptx.dml.color import RGBColor

from pptx.slide import Slide
from pptx.util import Inches, Pt


class Color(Enum):
    WHITE = RGBColor(255, 255, 255)
    BLACK = RGBColor(0, 0, 0)
    RED = RGBColor(255, 0, 0)
    GREEN = RGBColor(0, 255, 0)
    BLUE = RGBColor(0, 0, 255)


def setBackgroundImage(root: Presentation, slide: Slide, image_path: str):
    left = top = Inches(0)
    img_path = image_path
    pic = slide.shapes.add_picture(img_path, left, top, width=root.slide_width, height=root.slide_height)

    # Move it to the background
    slide.shapes._spTree.remove(pic._element)
    slide.shapes._spTree.insert(2, pic._element)


def setTitle(slide: Slide, text: str, color: Color = Color.BLACK, fontSize: int = 32):
    title = slide.shapes.title
    title.text = text
    title.text_frame.paragraphs[0].font.size = Pt(fontSize)
    title.text_frame.paragraphs[0].font.color.rgb = color.value


def addList(slide: Slide, text: List[str], color: Color = Color.BLACK, fontSize: int = 12):
    shapes = slide.shapes
    body_shape = shapes.placeholders[1]
    tf = body_shape.text_frame

    if len(text) == 0:
        return

    tf = body_shape.text_frame
    tf.text = text[0]

    for i in range(1, len(text)):
        p = tf.add_paragraph()
        p.text = text[i]

    for paragraph in body_shape.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(fontSize)
            run.font.color.rgb = color.value


def addTable(slide, data, color: Color = Color.BLACK, fontSize: int = 16):
    x, y, cx, cy = Inches(.25), Inches(3.5), Inches(9.5), Inches(1.5)
    shape = slide.shapes.add_table(len(data), len(data[0]), x, y, cx, cy)
    table = shape.table

    for i, row in enumerate(data):
        for j, cell in enumerate(row):
            table.cell(i, j).text = cell
            for paragraph in table.cell(i, j).text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(fontSize)
                    run.font.color.rgb = color.value


@click.command()
@click.option("--folder", "-f", help="Output folder to read from", required=True)
def main(folder: str):
    logging.info(f"Creating presentation from output folder: {folder}")

    root = Presentation()

    """ Ref for slide types: 
    0 ->  title and subtitle
    1 ->  title and content
    2 ->  section header
    3 ->  two content
    4 ->  Comparison
    5 ->  Title only 
    6 ->  Blank
    7 ->  Content with caption
    8 ->  Pic with caption
    """

    # Title Slide
    slide = root.slides.add_slide(root.slide_layouts[0])
    setBackgroundImage(root, slide, "bin/ppt-assets/background.jpg")
    setTitle(slide, f"{folder} Configuration Assessment Highlights", Color.WHITE, fontSize=48)
    info = json.loads(open(f"output/{folder}/info.json").read())
    slide.placeholders[1].text = f'Data As Of: {datetime.fromtimestamp(info["lastRun"], get_localzone()).strftime("%m-%d-%Y at %H:%M:%S")}'

    # Criteria Slide
    slide = root.slides.add_slide(root.slide_layouts[1])
    setTitle(slide, f"Configuration Assessment Tool Criteria")
    slide.shapes.add_picture("bin/ppt-assets/criteria.png", Inches(0.5), Inches(1.75), width=Inches(9), height=Inches(5))

    # Criteria Slide ctd...
    slide = root.slides.add_slide(root.slide_layouts[1])
    setTitle(slide, f"Configuration Assessment Tool Criteria ctd...")
    slide.shapes.add_picture("bin/ppt-assets/criteria2.png", Inches(0.5), Inches(1.75), width=Inches(9), height=Inches(4))

    # S/G/P Criteria & Scoring Slide
    slide = root.slides.add_slide(root.slide_layouts[1])
    setTitle(slide, f"B/S/G/P Criteria & Scoring")
    text = [
        "Platinum - 70% of criteria for an application must be Platinum, all remaining criteria must be at least Gold",
        "Gold - 70% of criteria for an application must be Gold or higher, all remaining criteria must be at least Silver",
        "Silver - 70% of criteria for an application must be Silver or higher",
        "Bronze - All remaining applications",
    ]
    addList(slide, text)

    wb = load_workbook(filename=f"output/{folder}/{folder}-MaturityAssessment-apm.xlsx")
    sheet = wb["Analysis"]
    scores = []
    for column_cell in sheet.iter_cols(1, sheet.max_column):
        if column_cell[0].value == 'OverallAssessment':
            j = 0
            for data in column_cell[1:]:
                scores.append(data.value)
            break


    data = [
        ["Controller", "Apps", "Bronze% (#)", "Silver% (#)", "Gold% (#)", "Platinum% (#)"],
        [
            folder,
            str(wb["Analysis"].max_row - 1),
            f"{format(scores.count('bronze')/(wb['Analysis'].max_row - 1) * 100, '.0f')}% ({scores.count('bronze')})",
            f"{format(scores.count('silver')/(wb['Analysis'].max_row - 1) * 100, '.0f')}% ({scores.count('silver')})",
            f"{format(scores.count('gold')/(wb['Analysis'].max_row - 1) * 100, '.0f')}% ({scores.count('gold')})",
            f"{format(scores.count('platinum')/(wb['Analysis'].max_row - 1) * 100, '.0f')}% ({scores.count('platinum')})",
        ],
    ]
    addTable(slide, data)

    # Saving file
    logging.info(f"Saving presentation to output/{folder}/{folder}-cx-presentation.pptx")
    root.save(f"output/{folder}/{folder}-cx-presentation.pptx")


if __name__ == "__main__":
    # cd to config-assessment-tool root directory
    path = os.path.realpath(f"{__file__}/../..")
    os.chdir(path)

    # create logs and output directories
    if not os.path.exists("logs"):
        os.makedirs("logs")
    if not os.path.exists("output"):
        os.makedirs("output")

    # init logging
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler("logs/asd.log"),
            logging.StreamHandler(),
        ],
    )

    logging.info(f"Working directory is {os.getcwd()}")

    main()
