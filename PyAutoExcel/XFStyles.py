from xlwt.Formatting import Alignment, Borders, Font, Pattern, Protection
from xlwt.Style import XFStyle as Style

from .ParserInterface2.Field import Field
from .ParserInterface2.Parser import Parser
from .ParserInterface2.Validators import TypeValidator


class XFFontConst:
    """
    Font constants for XF formatting.
    """

    ESCAPEMENT_NONE = 0x00
    ESCAPEMENT_SUPERSCRIPT = 0x01
    ESCAPEMENT_SUBSCRIPT = 0x02

    UNDERLINE_NONE = 0x00
    UNDERLINE_SINGLE = 0x01
    UNDERLINE_SINGLE_ACC = 0x21
    UNDERLINE_DOUBLE = 0x02
    UNDERLINE_DOUBLE_ACC = 0x22

    FAMILY_NONE = 0x00
    FAMILY_ROMAN = 0x01
    FAMILY_SWISS = 0x02
    FAMILY_MODERN = 0x03
    FAMILY_SCRIPT = 0x04
    FAMILY_DECORATIVE = 0x05

    CHARSET_ANSI_LATIN = 0x00
    CHARSET_SYS_DEFAULT = 0x01
    CHARSET_SYMBOL = 0x02
    CHARSET_APPLE_ROMAN = 0x4D
    CHARSET_ANSI_JAP_SHIFT_JIS = 0x80
    CHARSET_ANSI_KOR_HANGUL = 0x81
    CHARSET_ANSI_KOR_JOHAB = 0x82
    CHARSET_ANSI_CHINESE_GBK = 0x86
    CHARSET_ANSI_CHINESE_BIG5 = 0x88
    CHARSET_ANSI_GREEK = 0xA1
    CHARSET_ANSI_TURKISH = 0xA2
    CHARSET_ANSI_VIETNAMESE = 0xA3
    CHARSET_ANSI_HEBREW = 0xB1
    CHARSET_ANSI_ARABIC = 0xB2
    CHARSET_ANSI_BALTIC = 0xBA
    CHARSET_ANSI_CYRILLIC = 0xCC
    CHARSET_ANSI_THAI = 0xDE
    CHARSET_ANSI_LATIN_II = 0xEE
    CHARSET_OEM_LATIN_I = 0xFF


_xf_font_fields = [
    Field("height", TypeValidator(int), default=0x00C8),
    Field("italic", TypeValidator(int, bool), default=False),
    Field("struck_out", TypeValidator(int, bool), default=False),
    Field("outline", TypeValidator(int, bool), default=False),
    Field("shadow", TypeValidator(int, bool), default=False),
    Field("colour_index", TypeValidator(int), default=0x7FF),
    Field("bold", TypeValidator(int, bool), default=False),
    Field("_weight", TypeValidator(int), default=0x0190),
    Field("escapement", TypeValidator(int), default=0x00),
    Field("underline", TypeValidator(int), default=0x00),
    Field("family", TypeValidator(int), default=0x00),
    Field("charset", TypeValidator(int), default=0x01),
    Field("name", TypeValidator(str), default="Arial"),
]


class __XFFontParser(Parser):
    def __init__(self, params: dict):
        super().__init__(params, Font, _xf_font_fields)


def XFFont(
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
) -> Font:
    """
    Create a font formatting object.

    :param height: Font height in twips (1/20 of a point).
    :type height: int
    :param italic: Whether the font is italic.
    :type italic: bool
    :param struck_out: Whether the font is struck out.
    :type struck_out: bool
    :param outline: Whether the font is outlined.
    :type outline: bool
    :param shadow: Whether the font has a shadow.
    :type shadow: bool
    :param colour_index: Colour index of the font.
    :type colour_index: int
    :param bold: Whether the font is bold.
    :type bold: bool
    :param weight: Weight of the font.
    :type weight: int
    :param escapement: Escapement of the font.
    :type escapement: int
    :param underline: Underline of the font.
    :type underline: int
    :param family: Family of the font.
    :type family: int
    :param charset: Charset of the font.
    :type charset: int
    :param name: Name of the font. (e.g. "Arial")
    :type name: str
    :return: Font formatting object.
    :rtype: xlwt.Formatting.Font
    """
    return __XFFontParser(
        dict(
            height=height,
            italic=italic,
            struck_out=struck_out,
            outline=outline,
            shadow=shadow,
            colour_index=colour_index,
            bold=bold,
            weight=weight,
            escapement=escapement,
            underline=underline,
            family=family,
            charset=charset,
            name=name,
        )
    ).parse()


class XFAlignmentConst:
    """
    Alignment constants for XF formatting.
    """

    HORZ_GENERAL = 0x00
    HORZ_LEFT = 0x01
    HORZ_CENTER = 0x02
    HORZ_RIGHT = 0x03
    HORZ_FILLED = 0x04
    HORZ_JUSTIFIED = 0x05  # BIFF4-BIFF8X
    HORZ_CENTER_ACROSS_SEL = 0x06  # Centred across selection (BIFF4-BIFF8X)
    HORZ_DISTRIBUTED = 0x07  # Distributed (BIFF8X)

    VERT_TOP = 0x00
    VERT_CENTER = 0x01
    VERT_BOTTOM = 0x02
    VERT_JUSTIFIED = 0x03  # Justified (BIFF5-BIFF8X)
    VERT_DISTRIBUTED = 0x04  # Distributed (BIFF8X)

    DIRECTION_GENERAL = 0x00  # BIFF8X
    DIRECTION_LR = 0x01
    DIRECTION_RL = 0x02

    ORIENTATION_NOT_ROTATED = 0x00
    ORIENTATION_STACKED = 0x01
    ORIENTATION_90_CC = 0x02
    ORIENTATION_90_CW = 0x03

    ROTATION_0_ANGLE = 0x00
    ROTATION_STACKED = 0xFF

    WRAP_AT_RIGHT = 0x01
    NOT_WRAP_AT_RIGHT = 0x00

    SHRINK_TO_FIT = 0x01
    NOT_SHRINK_TO_FIT = 0x00


_xf_alignment_fields = [
    Field("horz", TypeValidator(int), default=XFAlignmentConst.HORZ_GENERAL),
    Field("vert", TypeValidator(int), default=XFAlignmentConst.VERT_BOTTOM),
    Field("dire", TypeValidator(int), default=XFAlignmentConst.DIRECTION_GENERAL),
    Field("orie", TypeValidator(int), default=XFAlignmentConst.ORIENTATION_NOT_ROTATED),
    Field("rota", TypeValidator(int), default=XFAlignmentConst.ROTATION_0_ANGLE),
    Field("wrap", TypeValidator(int), default=XFAlignmentConst.NOT_WRAP_AT_RIGHT),
    Field("shri", TypeValidator(int), default=XFAlignmentConst.NOT_SHRINK_TO_FIT),
    Field("inde", TypeValidator(int), default=0),
    Field("merg", TypeValidator(int), default=0),
]


class __XFAlignmentParser(Parser):
    def __init__(self, params: dict):
        super().__init__(params, Alignment, _xf_alignment_fields)


def XFAlignment(
    horz=XFAlignmentConst.HORZ_GENERAL,
    vert=XFAlignmentConst.VERT_BOTTOM,
    dire=XFAlignmentConst.DIRECTION_GENERAL,
    orie=XFAlignmentConst.ORIENTATION_NOT_ROTATED,
    rota=XFAlignmentConst.ROTATION_0_ANGLE,
    wrap=XFAlignmentConst.NOT_WRAP_AT_RIGHT,
    shri=XFAlignmentConst.NOT_SHRINK_TO_FIT,
    inde=0,
    merg=0,
) -> Alignment:
    """
    Create a alignnemnt formatting object.

    :param horz: Horizontal alignment.
    :type horz: int
    :param vert: Vertical alignment.
    :type vert: int
    :param dire: Text direction.
    :type dire: int
    :param orie: Orientation.
    :type orie: int
    :param rota: Rotation.
    :type rota: int
    :param wrap: Wrap text.
    :type wrap: int
    :param shri: Shrink to fit.
    :type shri: int
    :param inde: Indent level.
    :type inde: int
    :param merg: Merge cells.
    :type merg: int

    :return: Alignment formatting object.
    :rtype: xlwt.Formatting.Alignment
    """
    return __XFAlignmentParser(
        dict(
            horz=horz,
            vert=vert,
            dire=dire,
            orie=orie,
            rota=rota,
            wrap=wrap,
            shri=shri,
            inde=inde,
            merg=merg,
        )
    ).parse()


class XFBordersConst:
    """
    Border style constants.
    """

    NO_LINE = 0x00
    THIN = 0x01
    MEDIUM = 0x02
    DASHED = 0x03
    DOTTED = 0x04
    THICK = 0x05
    DOUBLE = 0x06
    HAIR = 0x07
    # The following for BIFF8
    MEDIUM_DASHED = 0x08
    THIN_DASH_DOTTED = 0x09
    MEDIUM_DASH_DOTTED = 0x0A
    THIN_DASH_DOT_DOTTED = 0x0B
    MEDIUM_DASH_DOT_DOTTED = 0x0C
    SLANTED_MEDIUM_DASH_DOTTED = 0x0D

    NEED_DIAG1 = 0x01
    NEED_DIAG2 = 0x01
    NO_NEED_DIAG1 = 0x00
    NO_NEED_DIAG2 = 0x00


_xf_borders_fields = [
    Field("left", TypeValidator(int), default=0x00),
    Field("right", TypeValidator(int), default=0x00),
    Field("top", TypeValidator(int), default=0x00),
    Field("bottom", TypeValidator(int), default=0x00),
    Field("diag", TypeValidator(int), default=0x00),
    Field("left_colour", TypeValidator(int), default=0x40),
    Field("right_colour", TypeValidator(int), default=0x40),
    Field("top_colour", TypeValidator(int), default=0x40),
    Field("bottom_colour", TypeValidator(int), default=0x40),
    Field("diag_colour", TypeValidator(int), default=0x40),
    Field("need_diag1", TypeValidator(int), default=0x00),
    Field("need_diag2", TypeValidator(int), default=0x00),
]


class __XFBordersParser(Parser):
    def __init__(self, params: dict):
        super().__init__(params, Borders, _xf_borders_fields)


def XFBorders(
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
) -> Borders:
    """
    Borders formatting.

    :param left: Style of the border at the left side.
    :type left: int
    :param right: Style of the border at the right side.
    :type right: int
    :param top: Style of the border at the top.
    :type top: int
    :param bottom: Style of the border at the bottom.
    :type bottom: int
    :param diag: Style of the diagonal border.
    :type diag: int
    :param left_colour: Border colour on the left side.
    :type left_colour: int
    :param right_colour: Border colour on the right side.
    :type right_colour: int
    :param top_colour: Border colour on the top.
    :type top_colour: int
    :param bottom_colour: Border colour on the bottom.
    :type bottom_colour: int
    :param diag_colour: Border colour for the diagonal.
    :type diag_colour: int
    :param need_diag1: Whether to draw a diagonal border.
    :type need_diag1: bool
    :param need_diag2: Whether to draw a diagonal border.
    :type need_diag2: bool
    :return: Borders formatting.
    :rtype: xlwt.Formatting.Borders
    """
    return __XFBordersParser(
        dict(
            left=left,
            right=right,
            top=top,
            bottom=bottom,
            diag=diag,
            left_colour=left_colour,
            right_colour=right_colour,
            top_colour=top_colour,
            bottom_colour=bottom_colour,
            diag_colour=diag_colour,
            need_diag1=need_diag1,
            need_diag2=need_diag2,
        )
    ).parse()


class XFPatternConst:
    """
    Constant values for the pattern of the cell.
    """

    NO_PATTERN = 0x00
    SOLID_PATTERN = 0x01


_xf_pattern_fields = [
    Field("pattern", TypeValidator(int), default=0x00),
    Field("pattern_fore_colour", TypeValidator(int), default=0x40),
    Field("pattern_back_colour", TypeValidator(int), default=0x41),
]


class __XFPatternParser(Parser):
    def __init__(self, params: dict):
        super().__init__(params, Pattern, _xf_pattern_fields)


def XFPattern(
    pattern=XFPatternConst.NO_PATTERN,
    pattern_fore_colour=0x40,
    pattern_back_colour=0x41,
) -> Pattern:
    """
    Pattern formatting.

    :param pattern: The type of the pattern.
    :type pattern: int
    :param pattern_fore_colour: The fore colour of the pattern.
    :type pattern_fore_colour: int
    :param pattern_back_colour: The back colour of the pattern.
    :type pattern_back_colour: int
    :return: Pattern formatting.
    :rtype: xlwt.Formatting.Pattern
    """
    return __XFPatternParser(
        dict(
            pattern=pattern,
            pattern_fore_colour=pattern_fore_colour,
            pattern_back_colour=pattern_back_colour,
        )
    ).parse()


_xf_protection_fields = [
    Field("cell_locked", TypeValidator(int, bool), default=False),
    Field("formula_hidden", TypeValidator(int, bool), default=False),
]


class __XFProtectionParser(Parser):
    def __init__(self, params: dict):
        super().__init__(params, Protection, _xf_protection_fields)


def XFProtection(
    cell_locked=True,
    formula_hidden=False,
) -> Protection:
    """
    Protection formatting.

    :param cell_locked: Whether the cell is locked.
    :type cell_locked: int
    :param formula_hidden: Whether the formula is hidden.
    :type formula_hidden: int
    :return: Protection formatting.
    :rtype: xlwt.Formatting.Protection
    """
    return __XFProtectionParser(
        dict(
            cell_locked=cell_locked,
            formula_hidden=formula_hidden,
        )
    ).parse()


_xf_style_fields = [
    Field("num_format_str", TypeValidator(str), default="General"),
    Field("font", TypeValidator(Font), default=Font()),
    Field("alignment", TypeValidator(Alignment), default=Alignment()),
    Field("borders", TypeValidator(Borders), default=Borders()),
    Field("pattern", TypeValidator(Pattern), default=Pattern()),
    Field("protection", TypeValidator(Protection), default=Protection()),
]


class __XFStyleParser(Parser):
    def __init__(self, params: dict):
        super().__init__(params, Style, _xf_style_fields)


def XFStyle(
    num_format_str: str = "General",
    font: Font = Font(),
    alignment: Alignment = Alignment(),
    borders: Borders = Borders(),
    pattern: Pattern = Pattern(),
    protection: Protection = Protection(),
) -> Style:
    """
    Style formatting.

    :param num_format_str: Number format string.
    :type num_format_str: str
    :param font: Font formatting.
    :type font: xlwt.Formatting.Font
    :param alignment: Alignment formatting.
    :type alignment: xlwt.Formatting.Alignment
    :param borders: Borders formatting.
    :type borders: xlwt.Formatting.Borders
    :param pattern: Pattern formatting.
    :type pattern: xlwt.Formatting.Pattern
    :param protection: Protection formatting.
    :type protection: xlwt.Formatting.Protection
    :return: Style formatting.
    :rtype: xlwt.Style.XFStyle
    """
    return __XFStyleParser(
        dict(
            num_format_str=num_format_str,
            font=font,
            alignment=alignment,
            borders=borders,
            pattern=pattern,
            protection=protection,
        )
    ).parse()
