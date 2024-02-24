from xlwt import ExcelFormula


def XLSHyperlink(url: str, desc: str = "") -> ExcelFormula.Formula:
    """
    It takes a URL and a description, and returns an Excel formula that will create a hyperlink in the cell

    :param url: The URL to link to
    :type url: str
    :param desc: The text to display in the cell
    :type desc: str
    :return: A hyperlink object
    """
    return ExcelFormula.Formula('HYPERLINK("{}";"{}")'.format(url, desc))
