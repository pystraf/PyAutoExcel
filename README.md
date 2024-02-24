# PyAutoExcel

An operation toolkit for MS Excel files.

## Usage

### I. Read Excel File

```python
from PyAutoExcel import ExcelReader
filename = "example.xlsx"  # The file to be read.
reader = ExcelReader(filename)
for sheet in reader.sheets():
    print(f"Sheet Name: {sheet.name}")
    for row in sheet.data: print(','.join(row))
    print()
```

### II. Write to Excel

```python
from PyAutoExcel import ExcelWriter, Sheet
data = [
    ['Name', 'Age', 'Sex'],
    ['Jenny', 18, 'Female'],
    ['Joe', 15, 'Male'],
    ['Jack', 8, 'Male'],
]
writer = ExcelWriter()

sheet = Sheet("Sheet1")   # Create worksheet
for i, row in enumerate(data): sheet.set_row(i, row)
writer.add_sheet(sheet)

writer.save("example.xlsx")
```

