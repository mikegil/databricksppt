#!python

import sys
import os

import click
import numpy as np
import pandas as pd
from pathlib import Path
from pptx import Presentation

from .databricksppt import toPPT, CHART_TYPE, LEGEND_POSITION


@click.command()
@click.argument('inputfile', type=click.Path(exists=True, dir_okay=False, resolve_path=True))
@click.argument('outputfile', type=click.Path())
@click.option('--inputfile2', type=click.Path(exists=True, dir_okay=False, resolve_path=True), help='Optional second data input file')
@click.option('--template', type=click.Path(exists=True, dir_okay=False, resolve_path=True), help='Create PPTX from given template')
@click.option('--layout-num', type=int, default=1, help='Layout # within template for new slide')
@click.option('--title', type=str, help='Title for slide on which to place data')
@click.option('--chart-title', type=str, help='Title for chart')
@click.option('--slide-num', type=int, default=0, help='Slide # on which to place data; 0 = new slide (default)')
@click.option('--placeholder-num', type=int, default=0, help='Placeholder # within the slide on which to place data')
@click.option('--chart-num', type=int, default=1, help='Chart # within the slide on which to place data - if not using Placeholders')
@click.option('--chart-type', type=click.Choice(list(map(lambda x: str(x.value), CHART_TYPE)), case_sensitive=False), default='Table', help='Type of chart to display (default = Table)')
@click.option('--legend-position', type=click.Choice(list(map(lambda x: str(x.value), LEGEND_POSITION)), case_sensitive=False), default='None', help='Position to display legend for chart (default = None)')
@click.option('--overlay-legend', is_flag=True, help='Places the legend in the chart area (overlaying the chart)')
@click.option('--column-names-as-labels', type=click.Choice(['True', 'False', 'Infer'], case_sensitive=False), default='Infer', help='Use DataFrame column names as series labels (default = Infer)')
@click.option('--first-column-as-labels', type=click.Choice(['True', 'False', 'Infer'], case_sensitive=False), default='Infer', help='Use values in first column as category labels(default=Infer)')
@click.option('--transpose', is_flag=True, help='Switches the rows from the dataframe to be categories and the columns to be series')
@click.option('--open', is_flag=True, help='Attempt to automatically open the PPTX file on success')
def main(inputfile, inputfile2, outputfile, template, layout_num, title, chart_title, slide_num, placeholder_num, chart_num, column_names_as_labels, first_column_as_labels, chart_type, legend_position, overlay_legend, transpose, open):
    """
    Runs databricksppt from the command line, using CSV input to produce a Powerpoint
    file including a Chart or Table built from this data
    """
    if (Path(outputfile).suffix != '.pptx'):
        outputfile += '.pptx'

    df = pd.read_csv(inputfile)  # , header=None)
    #df.name = "MyData"
    if (inputfile2 is not None):
        df2 = pd.read_csv(inputfile2)
        df = [df, df2]

    column_names_as_labels = None if column_names_as_labels == 'Infer' else True if column_names_as_labels == 'True' else False
    first_column_as_labels = None if first_column_as_labels == 'Infer' else True if first_column_as_labels == 'True' else False

    body_font = dict(
        name='Verdana',
        size=10
    )

    chart = dict(
        title=chart_title,
        placeholder_num=placeholder_num,
        chart_num=chart_num,
        chart_type=chart_type,
        legend_position=legend_position,
        overlay_legend=overlay_legend,
        data=df,
        column_names_as_labels=column_names_as_labels,
        first_column_as_labels=first_column_as_labels,
        transpose=transpose
    )

    slide = dict(
        layout_num=layout_num,
        title=title,
        slide_num=slide_num,
        charts=[chart]
    )

    presentation = dict(
        template=template,
        #      body_font=body_font,
        slides=[slide]
    )

    ppt = toPPT(presentation)
    if (isinstance(ppt, str)):
        print(ppt)
    else:
        ppt.save(outputfile)
        if open:
            os.system('open '+outputfile)
