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


def toPPT(slideInfo, chartInfo):
    pres = __create_presentation(slideInfo)
    if pres is None:
        return None

    slide = __create_slide(pres, slideInfo)
    if slide is None:
        return None

    placeholderNum = slideInfo.get('placeholder')
    if placeholderNum is not None:
        placeholder = __get_placeholder(slide, placeholderNum)
    else:
        chartNum = slideInfo.get('chart', 1)
        placeholder = __get_chart(slide, chartNum)

    if placeholder is None:
        return None
    __insert_object(slide, placeholder, chartInfo)

    return pres


def __create_presentation(slideInfo):
    template = slideInfo.get('template')
    if (template is not None):
        if (not isinstance(template, str)):
            template = None
        else:
            if (not path.isfile(template)):
                template = None

    return Presentation(template)


def __create_slide(pres, slideInfo):
    slideNum = slideInfo.get('slideNum', 0)
    layout = slideInfo.get('layout', 1)
    title = slideInfo.get('title')

    if slideNum == 0:
        slide = pres.slides.add_slide(pres.slide_layouts[layout])
    else:
        if len(pres.slides) >= slideNum:
            slide = pres.slides[slideNum-1]
        else:
            return None

    if slide.shapes.title is not None:
        slide.shapes.title.text = title

    return slide


def __get_placeholder(slide, placeholderNum):
    if len(slide.placeholders) <= placeholderNum:
        return None

    placeholderIdx = []

    for shape in slide.placeholders:
        placeholderIdx.append(shape.placeholder_format.idx)

    placeholder = slide.placeholders[placeholderIdx[placeholderNum]]

    # Remove empty placeholder
    sp = placeholder._sp
    sp.getparent().remove(sp)

    return placeholder


def __get_chart(slide, chartNum):
    if chartNum == 0:
        return None

    chartFound = 0

    for shape in slide.shapes:
        if shape.has_chart:
            chartFound += 1
        if chartFound == chartNum:
            shape.element.getparent().remove(shape.element)
            return shape

    return None


def __infer_category_labels(data):
    labelsInFirstCol = False

    for dataframe in data:
        firstCol = dataframe.iloc[:, 0]
        for cell in firstCol:
            if not isinstance(cell, numbers.Number):
                labelsInFirstCol = True

    return labelsInFirstCol


def __infer_series_labels(data):
    labelsInColumnHeaders = False

    for dataframe in data:
        for col in dataframe.columns:
            if not isinstance(col, numbers.Number):
                labelsInColumnHeaders = True

    return labelsInColumnHeaders


def __transpose_data(chartInfo):
    transposed_data = []

    for dataframe in chartInfo['data']:
        if not isinstance(dataframe, pd.DataFrame):
            return chartInfo

        if chartInfo['first_column_as_labels'] and chartInfo['column_names_as_labels']:
            indexColName = dataframe.columns[0]
            df = dataframe.set_index(
                dataframe.columns[0]).transpose().reset_index()
            df.rename(columns={'index': indexColName}, inplace=True)
        elif chartInfo['column_names_as_labels']:
            df = dataframe.transpose().reset_index()
        elif chartInfo['first_column_as_labels']:
            df = dataframe.set_index(dataframe.columns[0]).transpose()
        else:
            df = dataframe.transpose()

        transposed_data.append(df)

    chartInfo['data'] = transposed_data

    temp = chartInfo['column_names_as_labels']
    chartInfo['column_names_as_labels'] = chartInfo['first_column_as_labels']
    chartInfo['first_column_as_labels'] = temp

    return chartInfo


def __get_dataframes(data):
    if (not isinstance(data, pd.DataFrame) and not __iterable(data)):
        return None

    if (isinstance(data, pd.DataFrame)):
        dfs = [data]
    else:
        for dataframe in data:
            if not isinstance(dataframe, pd.DataFrame):
                return None
        dfs = data

    return dfs


def __insert_object(slide, placeholder, chartInfo):

    data = chartInfo.get('data')

    if (data is None):
        return

    if (isinstance(data, pd.DataFrame)):
        chartInfo['data'] = [data]

    for dataframe in chartInfo['data']:
        if not isinstance(dataframe, pd.DataFrame):
            return

    if not isinstance(chartInfo.get('column_names_as_labels'), bool):
        chartInfo['column_names_as_labels'] = __infer_series_labels(
            chartInfo['data'])

    if not isinstance(chartInfo.get('first_column_as_labels'), bool):
        chartInfo['first_column_as_labels'] = __infer_category_labels(
            chartInfo['data'])

    transpose = chartInfo.get('transpose', False)
    if transpose:
        chartInfo = __transpose_data(chartInfo)

    data = __get_dataframes(chartInfo.get('data'))
    dataframe = data[0]

    chart_type = chartInfo.get('chart_type', 'Table')

    if chart_type == CHART_TYPE.AREA.value:
        __insert_chart(XL_CHART_TYPE.AREA, slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.AREA_STACKED.value:
        __insert_chart(XL_CHART_TYPE.AREA_STACKED,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.AREA_STACKED_100.value:
        __insert_chart(XL_CHART_TYPE.AREA_STACKED_100,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.BAR.value:
        __insert_chart(XL_CHART_TYPE.BAR_CLUSTERED,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.BAR_STACKED.value:
        __insert_chart(XL_CHART_TYPE.BAR_STACKED,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.BAR_STACKED_100.value:
        __insert_chart(XL_CHART_TYPE.BAR_STACKED_100,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.COLUMN.value:
        __insert_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.COLUMN_STACKED.value:
        __insert_chart(XL_CHART_TYPE.COLUMN_STACKED,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.COLUMN_STACKED_100.value:
        __insert_chart(XL_CHART_TYPE.COLUMN_STACKED_100,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.LINE.value:
        __insert_chart(XL_CHART_TYPE.LINE, slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.LINE_STACKED.value:
        __insert_chart(XL_CHART_TYPE.LINE_STACKED,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.LINE_STACKED_100.value:
        __insert_chart(XL_CHART_TYPE.LINE_STACKED_100,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.LINE_MARKED.value:
        __insert_chart(XL_CHART_TYPE.LINE_MARKERS,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.LINE_MARKED_STACKED.value:
        __insert_chart(XL_CHART_TYPE.LINE_MARKERS_STACKED,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.LINE_MARKED_STACKED_100.value:
        __insert_chart(XL_CHART_TYPE.LINE_MARKERS_STACKED_100,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.DOUGHNUT.value:
        __insert_chart(XL_CHART_TYPE.DOUGHNUT, slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.DOUGHNUT_EXPLODED.value:
        __insert_chart(XL_CHART_TYPE.DOUGHNUT_EXPLODED,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.PIE.value:
        __insert_chart(XL_CHART_TYPE.PIE, slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.PIE_EXPLODED.value:
        __insert_chart(XL_CHART_TYPE.PIE_EXPLODED,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.RADAR.value:
        __insert_chart(XL_CHART_TYPE.RADAR, slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.RADAR_FILLED.value:
        __insert_chart(XL_CHART_TYPE.RADAR_FILLED,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.RADAR_MARKED.value:
        __insert_chart(XL_CHART_TYPE.RADAR_MARKERS,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.XY_SCATTER.value:
        __insert_xyzchart(XL_CHART_TYPE.XY_SCATTER,
                          slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.XY_SCATTER_LINES.value:
        __insert_xyzchart(XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
                          slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.XY_SCATTER_LINES_SMOOTHED.value:
        __insert_xyzchart(XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS,
                          slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.XY_SCATTER_LINES_MARKED.value:
        __insert_chart(XL_CHART_TYPE.XY_SCATTER_LINES,
                       slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.XY_SCATTER_LINES_MARKED_SMOOTHED.value:
        __insert_xyzchart(XL_CHART_TYPE.XY_SCATTER_SMOOTH,
                          slide, placeholder, chartInfo)
    elif chart_type == CHART_TYPE.BUBBLE.value:
        __insert_xyzchart(XL_CHART_TYPE.BUBBLE, slide, placeholder, chartInfo)
    else:
        __insert_table(slide, placeholder, chartInfo)


def __insert_table(slide, placeholder, chartInfo):
    df = chartInfo['data'][0]

    columns = df.shape[1]
    rows = df.shape[0]
    if chartInfo['column_names_as_labels']:
        rows += 1

    # Create new element with same shape and position as placeholder
    table = slide.shapes.add_table(
        rows, columns, placeholder.left, placeholder.top, placeholder.width, placeholder.height).table
    table.first_row = chartInfo['column_names_as_labels']
    table.first_col = chartInfo['first_column_as_labels']

    # Remove empty placeholder
    sp = placeholder._sp
    sp.getparent().remove(sp)

    # Populate table
    colNames = df.columns.tolist()

    rowNum = 0

    if chartInfo['column_names_as_labels']:
        col = 0
        for colName in colNames:
            table.cell(0, col).text = str(colName)
            col += 1
        rowNum += 1

    for index, row in df.iterrows():
        col = 0
        for colName in colNames:
            table.cell(rowNum, col).text = str(row.iloc[col])
            col += 1
        rowNum += 1

    return


def __iterable(obj):
    return isinstance(obj, Iterable)


def __create_chartdata(chartInfo):
    chart_data = CategoryChartData()

    # TODO: Deal with First Row as Labels and Column Names as Labels

    colNames = chartInfo['data'][0].columns.tolist()
    offset = 0

    if (chartInfo['first_column_as_labels']):
        offset = 1

    if len(colNames) > offset:

        colNum = 1
        for colName in colNames[offset:]:
            if (chartInfo['column_names_as_labels']):
                chart_data.categories.add_category(colName)
            else:
                chart_data.categories.add_category('Category '+str(colNum))

        rowNum = 1
        for index, row in chartInfo['data'][0].iterrows():
            data = []
            for colName in colNames[offset:]:
                data.append(row[colName])

            if chartInfo['first_column_as_labels']:
                chart_data.add_series(str(row[0]), data)
            else:
                chart_data.add_series('Series ' + str(rowNum), data)

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


def __insert_chart(chart_type, slide, placeholder, chartInfo):
    chart_data = __create_chartdata(chartInfo)
    if chart_data is None:
        return

    # Create new element with same shape and position as placeholder
    chart = slide.shapes.add_chart(chart_type, placeholder.left,
                                   placeholder.top, placeholder.width, placeholder.height, chart_data).chart

    __set_chart_title(chart, chartInfo)

    return


def __set_chart_title(chart, chartInfo):
    title = chartInfo.get('title')
    if title is not None:
        title_tf = chart.chart_title.text_frame
        title_tf.clear()
        title_p = title_tf.paragraphs[0]
        title_p.add_run().text = title


def __insert_xyzchart(chart_type, slide, placeholder, chartInfo):
    chart_data = __create_xyzdata(chartInfo['data'])
    if chart_data is None:
        return

    # Create new element with same shape and position as placeholder
    chart = slide.shapes.add_chart(chart_type, placeholder.left,
                                   placeholder.top, placeholder.width, placeholder.height, chart_data).chart

    __set_chart_title(chart, chartInfo)

    # Remove empty placeholder
    # sp = placeholder._sp
    # sp.getparent().remove(sp)

    return


def __get_datafile_name(filename):
    """
    return the default template file that comes with the package
    """
    return Path(__file__).parent / "data/" + filename
