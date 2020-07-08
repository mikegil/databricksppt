#!/usr/bin/env python

import sys
import os
from databricksppt import databricksppt

help = """
databricksppt Script

databricksppt file_to_process [output_file_name]

Creates databricksppt file from data
"""


def main():
    """
    startup function for running databricksppt as a script
    """
    try:
        infilename = sys.argv[1]
    except IndexError:
        print("you need to pass in a file name to process")
        print(help)
        sys.exit()
    try:
        outfilename = sys.argv[2]
    except IndexError:
        root, ext = os.path.splitext(infilename)
        outfilename = "test.html"

    # do the real work:
    print("Producing databricksppt from: %s and storing it in %s" %
          (infilename, outfilename))
    print(databricksppt.todatabricksppt(""))  # infilename, outfilename)

    print("done")
