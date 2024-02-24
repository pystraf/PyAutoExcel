"""
Create image objects for writing Excel.
"""
from os.path import isfile, join

import filetype
from openpyxl.drawing import image
from PIL import Image

from .ExcelTempFile import TEMPDIR


def _get_name():
    """
    Generate a unique bitmap file name within the TEMPDIR directory.

    :return: The unique bitmap file name.
    :rtype: str
    """
    DIR = join(TEMPDIR, "bitmaps")
    i = 0
    while True:
        p = join(DIR, "bitmap_") + str(i) + ".bmp"
        if not isfile(p):
            return p
        i += 1


def create_image_xls(file_name: str):
    """
    Create a bitmap image for writing .xls files.

    :param file_name: The name of the input file.
    :type file_name: str
    :return: The file name of the created or existing BMP image.
    :rtype: str
    """
    fp = open(file_name, "r")
    fmt = filetype.guess(file_name).extension
    if fmt == "bmp":
        fp.close()
        return file_name
    else:
        img = Image.open(fp)
        fp.close()
        target_path = _get_name()
        fw = open(target_path, "w")
        img.save(fw, "bmp")
        fw.close()
        return file_name


def create_image_xlsx(file_name: str):
    """
    Create an image for writing .xlsx files.

    :param file_name: The name of the image.
    :type file_name: str
    :return: The created image object
    :rtype: openpyxl.drawing.image.Image
    """
    with open(file_name, "rb") as fp:
        img = image.Image(img=Image.open(fp))
    return img
