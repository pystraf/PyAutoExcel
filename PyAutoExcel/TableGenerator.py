"""
Creating Tables in HTML with data.
"""


class BasicTableGenerator:
    """
    A basic table generator for HTML.

    :param table_data: The data for the table.
    :type table_data: list[list]
    """

    def __init__(self, table_data: list[list]):
        """
        Initialize the BasicTableGenerator with table data.

        :param table_data: The data for the table.
        :type table_data: list[list]
        """
        self.table_data = table_data
        self.content = (
            "<!--generate by PyAutoExcel.TableGeneator.BasicTableGenerator-->\n"
        )

    def start(self):
        """
        Start generating the table.
        """
        self.content += "<table>\n"

    def start_row(self):
        """
        Start a new row in the table.
        """
        self.content += "    <tr>\n"

    def cell(self, value):
        """
        Add a cell with the given value to the current row.

        :param value: The value for the cell.
        :type value: any
        """
        self.content += f"        <td>{value}</td>\n"

    def end_row(self):
        """
        End the current row in the table.
        """
        self.content += "    </tr>\n"

    def end(self):
        """
        End generating the table.
        """
        self.content += "</table>\n"

    def generate(self):
        """
        Generate the table based on the provided data.

        :return: The generated table.
        :rtype: BasicTableGenerator
        """
        self.start()
        for row in self.table_data:
            self.start_row()
            for cell in row:
                self.cell(value=cell)
            self.end_row()
        self.end()
        return self


class CustomTableGenerator(BasicTableGenerator):
    """
    A custom table generator class that extends the BasicTableGenerator class.
    """

    def __init__(
        self,
        table_data: list[list],
        table_option: str = "",
        row_option: str = "",
        cell_option: str = "",
    ):
        super().__init__(table_data)
        self.table_option = table_option
        self.row_option = row_option
        self.cell_option = cell_option
        self.content = (
            "<!--generate by PyAutoExcel.TableGeneator.CustomTableGenerator-->\n"
        )

    def start(self):
        """
        Override the start method to add custom logic.
        """
        self.content += f"<table {self.table_option}>\n"

    def start_row(self):
        """
        Override the start_row method to add custom logic.
        """
        self.content += f"    <tr {self.row_option}>\n"

    def cell(self, value):
        """
        Override the cell method to add custom logic.
        """
        self.content += f"        <td {self.cell_option}>{value}</td>\n"
