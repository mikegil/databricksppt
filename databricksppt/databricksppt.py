from pathlib import Path
from os import path
from enum import Enum
import numbers
from collections.abc import Iterable
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData, XyChartData, BubbleChartData
from pptx.enum.shapes import PP_PLACEHOLDER
from itertools import islice
import pandas as pd
import numpy as np


class CHART_TYPE(Enum):
    AREA = 'Area'
    AREA_STACKED = 'Area-Stacked'
    AREA_STACKED_100 = 'Area-Stacked-100'
    BAR = 'Bar'
    BAR_STACKED = 'Bar-Stacked'
    BAR_STACKED_100 = 'Bar-Stacked-100'
    COLUMN = 'Column'
    COLUMN_STACKED = 'Column-Stacked'
    COLUMN_STACKED_100 = 'Column-Stacked-100'
    LINE = 'Line'
    LINE_STACKED = 'Line-Stacked'
    LINE_STACKED_100 = 'Line-Stacked-100'
    LINE_MARKED = 'Line-Marked'
    LINE_MARKED_STACKED = 'Line-Marked-Stacked'
    LINE_MARKED_STACKED_100 = 'Line-Marked-Stacked-100'
    DOUGHNUT = 'Doughnut'
    DOUGHNUT_EXPLODED = 'Doughnut-Exploded'
    PIE = 'Pie'
    PIE_EXPLODED = 'Pie-Exploded'
    RADAR = 'Radar'
    RADAR_FILLED = 'Radar-Filled'
    RADAR_MARKED = 'Radar-Marked'
    XY_SCATTER = 'XY-Scatter'
    XY_SCATTER_LINES = 'XY-Scatter-Lines'
    XY_SCATTER_LINES_SMOOTHED = 'XY-Scatter-Lines-Smoothed'
    XY_SCATTER_LINES_MARKED = 'XY-Scatter-Lines-Marked'
    XY_SCATTER_LINES_MARKED_SMOOTHED = 'XY-Scatter-Lines-Marked-Smoothed'
    BUBBLE = 'Bubble'
    TABLE = 'Table'


def chart_types():
    values = []
    for chart_type in CHART_TYPE:
        values.append(chart_type.value)
    return values


def toPPT(dfs, template=None, layout=1, title=None, subtitle="", slideNum=0, chart_type='Table', transpose=False):
    if (not isinstance(dfs, pd.DataFrame) and not __iterable(dfs)):
        return None

    if (isinstance(dfs, pd.DataFrame)):
        dfs = [dfs]
    else:
        for dataframe in dfs:
            if not isinstance(dataframe, pd.DataFrame):
                return None

    if transpose:
        dfs = __transpose_dfs(dfs)

    df = dfs[0]

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

    placeholder = slide.placeholders[placeholderIdx[0]]
    phf = placeholder.placeholder_format
    if phf.type == PP_PLACEHOLDER.TITLE:
        if len(slide.placeholders) < 2:
            return
        else:
            __set_titles(placeholder, title, subtitle)
            placeholder = slide.placeholders[placeholderIdx[1]]

    if chart_type == CHART_TYPE.AREA.value:
        __insert_chart(XL_CHART_TYPE.AREA, slide, placeholder, df)
    elif chart_type == CHART_TYPE.AREA_STACKED.value:
        __insert_chart(XL_CHART_TYPE.AREA_STACKED, slide, placeholder, df)
    elif chart_type == CHART_TYPE.AREA_STACKED_100.value:
        __insert_chart(XL_CHART_TYPE.AREA_STACKED_100, slide, placeholder, df)
    elif chart_type == CHART_TYPE.BAR.value:
        __insert_chart(XL_CHART_TYPE.BAR_CLUSTERED, slide, placeholder, df)
    elif chart_type == CHART_TYPE.BAR_STACKED.value:
        __insert_chart(XL_CHART_TYPE.BAR_STACKED, slide, placeholder, df)
    elif chart_type == CHART_TYPE.BAR_STACKED_100.value:
        __insert_chart(XL_CHART_TYPE.BAR_STACKED_100, slide, placeholder, df)
    elif chart_type == CHART_TYPE.COLUMN.value:
        __insert_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, slide, placeholder, df)
    elif chart_type == CHART_TYPE.COLUMN_STACKED.value:
        __insert_chart(XL_CHART_TYPE.COLUMN_STACKED, slide, placeholder, df)
    elif chart_type == CHART_TYPE.COLUMN_STACKED_100.value:
        __insert_chart(XL_CHART_TYPE.COLUMN_STACKED_100,
                       slide, placeholder, df)
    elif chart_type == CHART_TYPE.LINE.value:
        __insert_chart(XL_CHART_TYPE.LINE, slide, placeholder, df)
    elif chart_type == CHART_TYPE.LINE_STACKED.value:
        __insert_chart(XL_CHART_TYPE.LINE_STACKED, slide, placeholder, df)
    elif chart_type == CHART_TYPE.LINE_STACKED_100.value:
        __insert_chart(XL_CHART_TYPE.LINE_STACKED_100, slide, placeholder, df)
    elif chart_type == CHART_TYPE.LINE_MARKED.value:
        __insert_chart(XL_CHART_TYPE.LINE_MARKERS, slide, placeholder, df)
    elif chart_type == CHART_TYPE.LINE_MARKED_STACKED.value:
        __insert_chart(XL_CHART_TYPE.LINE_MARKERS_STACKED,
                       slide, placeholder, df)
    elif chart_type == CHART_TYPE.LINE_MARKED_STACKED_100.value:
        __insert_chart(XL_CHART_TYPE.LINE_MARKERS_STACKED_100,
                       slide, placeholder, df)
    elif chart_type == CHART_TYPE.DOUGHNUT.value:
        __insert_chart(XL_CHART_TYPE.DOUGHNUT, slide, placeholder, df)
    elif chart_type == CHART_TYPE.DOUGHNUT_EXPLODED.value:
        __insert_chart(XL_CHART_TYPE.DOUGHNUT_EXPLODED, slide, placeholder, df)
    elif chart_type == CHART_TYPE.PIE.value:
        __insert_chart(XL_CHART_TYPE.PIE, slide, placeholder, df)
    elif chart_type == CHART_TYPE.PIE_EXPLODED.value:
        __insert_chart(XL_CHART_TYPE.PIE_EXPLODED, slide, placeholder, df)
    elif chart_type == CHART_TYPE.RADAR.value:
        __insert_chart(XL_CHART_TYPE.RADAR, slide, placeholder, df)
    elif chart_type == CHART_TYPE.RADAR_FILLED.value:
        __insert_chart(XL_CHART_TYPE.RADAR_FILLED, slide, placeholder, df)
    elif chart_type == CHART_TYPE.RADAR_MARKED.value:
        __insert_chart(XL_CHART_TYPE.RADAR_MARKERS, slide, placeholder, df)
    elif chart_type == CHART_TYPE.XY_SCATTER.value:
        __insert_xyzchart(XL_CHART_TYPE.XY_SCATTER, slide, placeholder, dfs)
    elif chart_type == CHART_TYPE.XY_SCATTER_LINES.value:
        __insert_xyzchart(XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
                          slide, placeholder, dfs)
    elif chart_type == CHART_TYPE.XY_SCATTER_LINES_SMOOTHED.value:
        __insert_xyzchart(XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS,
                          slide, placeholder, dfs)
    elif chart_type == CHART_TYPE.XY_SCATTER_LINES_MARKED.value:
        __insert_chart(XL_CHART_TYPE.XY_SCATTER_LINES, slide, placeholder, dfs)
    elif chart_type == CHART_TYPE.XY_SCATTER_LINES_MARKED_SMOOTHED.value:
        __insert_xyzchart(XL_CHART_TYPE.XY_SCATTER_SMOOTH,
                          slide, placeholder, dfs)
    elif chart_type == CHART_TYPE.BUBBLE.value:
        __insert_xyzchart(XL_CHART_TYPE.BUBBLE, slide, placeholder, dfs)
    else:
        __insert_table(slide, placeholder, df)

    BUBBLE = 'Bubble'

    return pres


def __transpose_dfs(dfs):
    tdfs = []
    for dataframe in dfs:
        labelsInFirstCol = True
        firstCol = dataframe.iloc[:, 0]
        for cell in firstCol:
            if isinstance(cell, numbers.Number):
                labelsInFirstCol = False

        labelsInColumnHeaders = True
        for col in dataframe.columns:
            if isinstance(col, numbers.Number):
                labelsInColumnHeaders = False

        if labelsInFirstCol and labelsInColumnHeaders:
            df1 = dataframe.set_index(
                dataframe.columns[0]).transpose().reset_index()
        elif labelsInColumnHeaders:
            df1 = dataframe.transpose().reset_index()
        else:
            df1 = dataframe.set_index(dataframe.columns[0]).transpose()

        tdfs.append(df1)
#            tdfs.append(dataframe.set_index(dataframe.columns[0]).transpose())

    return tdfs


def __insert_table(slide, placeholder, df):
    colNames = df.columns.tolist()

    # Create new element with same shape and position as placeholder
    table = slide.shapes.add_table(
        df.shape[0]+1, df.shape[1], placeholder.left, placeholder.top, placeholder.width, placeholder.height).table

    # Remove empty placeholder
    sp = placeholder._sp
    sp.getparent().remove(sp)

    # Populate table
    col = 0
    for colName in colNames:
        table.cell(0, col).text = str(colName)
        col += 1

    rowNum = 1
    for index, row in df.iterrows():
        col = 0
        for colName in colNames:
            table.cell(rowNum, col).text = str(row.iloc[col])
            col += 1
        rowNum += 1

    return


def __iterable(obj):
    return isinstance(obj, Iterable)


def __create_chartdata(df):
    chart_data = CategoryChartData()

    colNames = df.columns.tolist()

    if len(colNames) >= 2:

        for colName in colNames[1:]:
            chart_data.categories.add_category(colName)

        for index, row in df.iterrows():
            data = []
            for colName in colNames[1:]:
                data.append(row[colName])

            chart_data.add_series(str(row[0]), data)

    return chart_data


def __create_xyzdata(dfs):
    chart_data = None

    seriesNum = 1

    for df in dfs:
        colNames = df.columns.tolist()
        name = 'Series ' + str(seriesNum)
        if hasattr(df, 'name') and df.name != "":
            name = df.name

        if len(colNames) > 1 and len(colNames) < 4:
            if len(colNames) == 2 and chart_data is None:
                chart_data = XyChartData()
            elif len(colNames) == 3 and chart_data is None:
                chart_data = BubbleChartData()

            series = chart_data.add_series(name)
            for index, row in df.iterrows():
                data = []
                for colName in colNames:
                    data.append(row[colName])

                if len(colNames) == 2:
                    series.add_data_point(data[0], data[1])
                else:
                    series.add_data_point(data[0], data[1], data[2])

            seriesNum += 1

    return chart_data


def __insert_chart(chart_type, slide, placeholder, df):
    chart_data = __create_chartdata(df)

    # Create new element with same shape and position as placeholder
    chart = slide.shapes.add_chart(chart_type, placeholder.left,
                                   placeholder.top, placeholder.width, placeholder.height, chart_data).chart

    # Remove empty placeholder
    sp = placeholder._sp
    sp.getparent().remove(sp)

    return


def __insert_xyzchart(chart_type, slide, placeholder, dfs):
    chart_data = __create_xyzdata(dfs)

    # Create new element with same shape and position as placeholder
    chart = slide.shapes.add_chart(chart_type, placeholder.left,
                                   placeholder.top, placeholder.width, placeholder.height, chart_data).chart

    # Remove empty placeholder
    sp = placeholder._sp
    sp.getparent().remove(sp)

    return


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
