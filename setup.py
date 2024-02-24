from setuptools import setup

requires = [
    'xlrd>=2',
    'xlwt>=1.3',
    'xlutils>=2',
    'openpyxl>=3.0.10,<3.5,!=3.2.0b1',
    'xlsxwriter>=3.0.7',
    'xlsxlite>=0.2',
    'xlsxcessive>=1.1',
    'odswriter>=0.4',
    'python-xlsxio>=0.1.5',
    'pandas>=1.5,<3',
    'numpy>=1.24,<2',
    'prettytable>=3,<4',
    'proglog>=0.1.10,<1',
    'rich>=12',
]

setup(
    name='PyAutoExcel',
    version='3.0.2',
    python_requires='>=3.9',
    packages=['PyAutoExcel', 'PyAutoExcel.Tiny', 'PyAutoExcel.Command', 'PyAutoExcel.Engines',
              'PyAutoExcel.Documents', 'PyAutoExcel.Documents.File',
              'PyAutoExcel.Documents.File.Excel', 'PyAutoExcel.Documents.File.Excel.Reader',
              'PyAutoExcel.Documents.File.Excel.Writer', 'PyAutoExcel.Documents.Workbook',
              'PyAutoExcel.Documents.Workbook.BookImpl', 'PyAutoExcel.TempFiles'],
    url='',
    install_requires=requires,
    extras_require={'pdf': ['pdfkit'], 'image': ['imgkit']},
    license='MIT',
    author='pystraf',
    author_email='pystraf@163.com',
    description='An operation toolkit for MS Excel files.'
)
