"""
Extract the VBA project file from an .xlsxm file.

"""
import zipfile


class Extractor:
    """
    A class for extracting VBA project from an Excel file.
    """

    zip: zipfile.ZipFile

    def __init__(self, file_name: str):
        """
        Initialize the Extractor with the given file name.

        :param file_name: The name of the Excel file.
        :type file_name: str
        """
        self.file_name = file_name
        self.zip = None
        self.data = None

    def open(self, using_zip64: bool = True):
        """
        Open the Excel file as a zip archive.

        :param using_zip64: Whether to use ZIP64 format.
        :type using_zip64: bool
        """
        self.zip = zipfile.ZipFile(self.file_name, "r", allowZip64=using_zip64)

    def extract(self):
        """
        Extract the VBA project from the Excel file.
        """
        self.data = self.zip.read("xl/vbaProject.bin")

    def save(self, file_name: str):
        """
        Save the extracted VBA project to a file.

        :param file_name: The name of the file to save the VBA project to.
        :type file_name: str
        """
        with open(file_name, "wb") as fp:
            fp.write(self.data)


def extract_vba_project(
    xlsm_file_name: str,
    save_file_name: str = "vbaProject.bin",
    using_zip64: bool = True,
):
    """
    Extract the VBA project from an Excel file and save it to a file.

    :param xlsm_file_name: The name of the Excel file.
    :type xlsm_file_name: str
    :param save_file_name: The name of the file to save the VBA project to.
    :type save_file_name: str
    :param using_zip64: Whether to use ZIP64 format.
    :type using_zip64: bool
    """
    extractor = Extractor(file_name=xlsm_file_name)
    extractor.open(using_zip64=using_zip64)
    extractor.extract()
    extractor.save(file_name=save_file_name)
    extractor.zip.close()
