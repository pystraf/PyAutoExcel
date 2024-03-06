import xlwt

from .XFStyles import (XFAlignment, XFAlignmentConst, XFBorders,
                       XFBordersConst, XFFont, XFFontConst, XFPattern,
                       XFPatternConst, XFProtection, XFStyle)


def set_style(
    # font
    height=0x00C8,  # 200: this is font with height 10 points
    italic=False,
    struck_out=False,
    outline=False,
    shadow=False,
    colour_index=0x7FFF,
    bold=False,
    weight=0x0190,
    escapement=XFFontConst.ESCAPEMENT_NONE,
    underline=XFFontConst.UNDERLINE_NONE,
    family=XFFontConst.FAMILY_NONE,
    charset=XFFontConst.CHARSET_SYS_DEFAULT,
    name="Arial",
    # alignment
    horz=XFAlignmentConst.HORZ_GENERAL,
    vert=XFAlignmentConst.VERT_BOTTOM,
    dire=XFAlignmentConst.DIRECTION_GENERAL,
    orie=XFAlignmentConst.ORIENTATION_NOT_ROTATED,
    rota=XFAlignmentConst.ROTATION_0_ANGLE,
    wrap=XFAlignmentConst.NOT_WRAP_AT_RIGHT,
    shri=XFAlignmentConst.NOT_SHRINK_TO_FIT,
    inde=0,
    merg=0,
    # borders
    left=XFBordersConst.NO_LINE,
    right=XFBordersConst.NO_LINE,
    top=XFBordersConst.NO_LINE,
    bottom=XFBordersConst.NO_LINE,
    diag=XFBordersConst.NO_LINE,
    left_colour=0x40,
    right_colour=0x40,
    top_colour=0x40,
    bottom_colour=0x40,
    diag_colour=0x40,
    need_diag1=XFBordersConst.NO_NEED_DIAG1,
    need_diag2=XFBordersConst.NO_NEED_DIAG2,
    # pattern
    pattern=XFPatternConst.NO_PATTERN,
    pattern_fore_colour=0x40,
    pattern_back_colour=0x41,
    # protection
    cell_locked=1,
    formula_hidden=0,
    # other
    num_format_str="General",
) -> xlwt.XFStyle:
    font = XFFont(
        height,
        italic,
        struck_out,
        outline,
        shadow,
        colour_index,
        bold,
        weight,
        escapement,
        underline,
        family,
        charset,
        name,
    )
    alignment = XFAlignment(horz, vert, dire, orie, rota, wrap, shri, inde, merg)
    borders = XFBorders(
        left,
        right,
        top,
        bottom,
        diag,
        left_colour,
        right_colour,
        top_colour,
        bottom_colour,
        diag_colour,
        need_diag1,
        need_diag2,
    )
    pattern = XFPattern(pattern, pattern_fore_colour, pattern_back_colour)
    protection = XFProtection(cell_locked, formula_hidden)
    style = XFStyle(num_format_str, font, alignment, borders, pattern, protection)
    return style

