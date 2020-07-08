from pathlib import Path


def toPPT(df):
    lines = ""
    with open(get_templatefile_name()) as template:
        for line in template:
            lines += line
    return lines


def get_templatefile_name():
    """
    return the default template file that comes with the package
    """
    return Path(__file__).parent / "data/template.html"
