from pathlib import Path
from os import path
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import PP_PLACEHOLDER
import pandas as pd


def toPPT(df, template=None, layout=1, title=None, subtitle="", slideNum=0):
    if (not isinstance(df, pd.DataFrame)):
        return None

    if (template is not None):
        if (not isinstance(template, str)):
            template = None
        else:
            if (not template == "" and not path.isfile(template)):
                template = None

    pres = Presentation(template)
    slide = None

    if slideNum == 0:
        slide = pres.slides.add_slide(pres.slide_layouts[1])
    else:
        if len(pres.slides) >= slideNum:
            slide = pres.slides[slideNum-1]
        else:
            return None

    #placeholder = slide.shapes[1]

    if len(slide.placeholders) == 0:
        return None

    placeholderIdx = []

    for shape in slide.placeholders:
        placeholderIdx.append(shape.placeholder_format.idx)

    shape = slide.placeholders[placeholderIdx[0]]
    phf = shape.placeholder_format
    if phf.type == PP_PLACEHOLDER.TITLE:
        if len(slide.placeholders) < 2:
            return
        else:
            __set_titles(shape, title, subtitle)
            shape = slide.placeholders[placeholderIdx[1]]

    print('%d, %s' % (phf.idx, phf.type))
    print('%d, %d, %d, %d' % (shape.left, shape.top, shape.width, shape.height))

    x, y, cx, cy = Inches(2), Inches(2), Inches(4), Inches(1.5)
    colNames = df.columns.tolist()

    # table = placeholder.insert_table(cols=df.shape[0], rows=df.shape[1])

    # col = 0
    # for colName in colNames:
    #     table.cell(0, col).text = colName
    #     col += 1

    # for index, rows in df.iterrows():
    #     col = 0
    #     for colName in colNames:
    #         table.cell(index+1, col).text = colName
    #         col += 1

    return pres


def __set_titles(titleph, title, subtitle):
    if not titleph.has_text_frame:
        return
    if title is None or not isinstance(title, str):
        return

    text_frame = titleph.text_frame
    p = text_frame.paragraphs[0]
    p.text = title
    if (len(text_frame.paragraphs) > 1 and subtitle is not None and isinstance(subtitle, str)):
        p = text_frame.paragraphs[1]
        p.text = subtitle


def __get_datafile_name(filename):
    """
    return the default template file that comes with the package
    """
    return Path(__file__).parent / "data/" + filename
