#!python

import sys
import os
import click
import numpy as np
import pandas as pd
from pathlib import Path
from databricksppt import databricksppt
from pptx import Presentation


@click.command()
@click.argument('inputfile', type=click.Path(exists=True, dir_okay=False, resolve_path=True))
@click.argument('outputfile', type=click.Path())
@click.option('--template', type=click.Path(exists=True, dir_okay=False, resolve_path=True), help='Create PPTX from given template')
@click.option('--layout', type=int, help='Layout within template for new slide')
@click.option('--title', type=str, help='Title for slide on which to place data')
@click.option('--subtitle', type=str, help='Subtitle for slide on which to place data')
@click.option('--slide', type=int, default=0, help='Slide # on which to place data; 0 = new slide (default)')
@click.option('--chart-type', type=click.Choice(['Table', 'Bar'], case_sensitive=False), default='Table', help='Type of chart to display (default = Table)')
@click.option('--open', is_flag=True, help='Attempt to automatically open the PPTX file on success')
def main(inputfile, outputfile, template, layout, title, subtitle, slide, chart_type, open):
    """
    Runs databricksppt from the command line, using CSV input to produce a Powerpoint
    file including a Chart or Table built from this data
    """
    if (Path(outputfile).suffix != '.pptx'):
        outputfile += '.pptx'

    df = pd.read_csv(inputfile)
    pres = databricksppt.toPPT(
        df, template=template, layout=layout, title=title, subtitle=subtitle, slideNum=slide, chart_type=chart_type)
    if (pres is not None):
        pres.save(outputfile)
        if open:
            os.system('open '+outputfile)
    else:
        print("No PPT Created")
