"""
Convert the styles read in xlrd to the style class used by xlwt.
"""
import xlrd
from xlrd.formatting import (XF, Font, XFAlignment, XFBackground, XFBorder,
                             XFProtection)
from xlwt import XFStyle

from . import XFStyles


class FontBridge:
    """
    Bridge between xlrd Font and xlwt XFStyle.Font.
    """

    def __init__(self, font: Font):
        self.font = font

    def migrate(self):
        ft = self.font
        return XFStyles.XFFont(
            height=ft.height,
            italic=ft.italic,
            struck_out=ft.struck_out,
            outline=ft.outline,
            shadow=ft.shadow,
            colour_index=ft.colour_index,
            bold=ft.bold,
            weight=ft.weight,
            escapement=ft.escapement,
            underline=ft.underline_type,
            family=ft.family,
            charset=ft.character_set,
            name=ft.name,
        )


class ProtectionBridge:
    """
    Bridge between xlrd XFProtection and xlwt XFStyle.Protection.
    """

    def __init__(self, protection: XFProtection):
        self.protection = protection

    def migrate(self):
        pt = self.protection
        return XFStyles.XFProtection(
            cell_locked=pt.cell_locked,
            formula_hidden=pt.formula_hidden,
        )


# class BorderBridge:   # Which?
class BordersBridge:
    """
    Bridge between xlrd XFBorder and xlwt XFStyles.XFBorders.
    """

    def __init__(self, borders: XFBorder):
        self.borders = borders

    def migrate(self):
        bd = self.borders
        return XFStyles.XFBorders(
            left=bd.left_line_style,
            right=bd.right_line_style,
            top=bd.top_line_style,
            bottom=bd.bottom_line_style,
            diag=bd.diag_line_style,
            left_colour=bd.left_colour_index,
            right_colour=bd.right_colour_index,
            top_colour=bd.top_colour_index,
            bottom_colour=bd.bottom_colour_index,
            diag_colour=bd.diag_colour_index,
            need_diag1=bd.diag_down,
            need_diag2=bd.diag_up,
        )


# class BackgroundBridge:   # Which?
class PatternBridge:
    """
    Bridge between xlrd XFBackground and xlwt XFStyles.XFPattern.
    """

    def __init__(self, pattern: XFBackground):
        self.pattern = pattern

    def migrate(self):
        pn = self.pattern
        return XFStyles.XFPattern(
            pattern=pn.fill_pattern,
            pattern_fore_colour=pn.pattern_colour_index,
            pattern_back_colour=pn.background_colour_index,
        )


class AlignmentBridge:
    """
    Bridge between xlrd XFAlignment and xlwt XFStyles.XFAlignment.
    """

    def __init__(self, alignment: XFAlignment):
        self.alignment = alignment

    def migrate(self):
        am = self.alignment
        return XFStyles.XFAlignment(
            horz=am.hor_align,
            vert=am.vert_align,
            dire=am.text_direction,
            rota=am.rotation,
            wrap=am.text_wrapped,
            shri=am.shrink_to_fit,
            inde=am.indent_level,
        )


class StyleBridge:
    """
    Bridge between xlrd XF and xlwt XFStyles.XFStyle.
    """

    def __init__(self, style: XF, parent: xlrd.Book):
        self.style = style
        self.format_str = parent.format_map[style.format_key].format_str
        self.font = FontBridge(font=parent.font_list[style.font_index]).migrate()
        self.protection = ProtectionBridge(protection=style.protection).migrate()
        self.borders = BordersBridge(borders=style.border).migrate()
        self.pattern = PatternBridge(pattern=style.background).migrate()
        self.alignment = AlignmentBridge(alignment=style.alignment).migrate()

    def migrate(self):
        return XFStyles.XFStyle(
            num_format_str=self.format_str,
            font=self.font,
            alignment=self.alignment,
            borders=self.borders,
            pattern=self.pattern,
            protection=self.protection,
        )


def migrate_style(book: xlrd.Book) -> list[XFStyle]:
    """
    Migrate xlrd XF to xlwt XFStyles.XFStyle.

    :param book: xlrd workbook.
    :type book: xlrd.Book
    :return: list of XFStyles.XFStyle.
    :rtype: list[XFStyle]
    :raises IOError: if formatting_info=False.
    """
    if not book.formatting_info:
        raise IOError(
            "Unable in formatting_info=False. "
            "Please reopen the workbook with formatting_info=True."
        )
    result = []
    for xf in book.xf_list:
        style = StyleBridge(style=xf, parent=book).migrate()
        result.append(style)
    return result.copy()
