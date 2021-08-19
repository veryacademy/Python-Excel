# **001-Reading Excel Files with openpyxl**

Welcome to the Python Excel Series.

**Video Tutorial:** 

#### **Difficulty:** Beginner
#### **Tags/Keywords:** Python, Excel, openpyxl
---

### Code Examples

```python
# ex1.0.1 Loading the Workbook

from openpyxl import load_workbook

# https://www.pythonexcel.com/openpyxl-load-workbook-function.php
# https://openpyxl.readthedocs.io/en/stable/optimized.html?highlight=load_workbook#read-only-mode
####
# read_only=True 
# write_only=True
# data_only 

wb = load_workbook(filename = 'example.xlsx')
print(wb)
```
```python
# ex1.0.2 Working with Sheets

from openpyxl import load_workbook

wb = load_workbook(filename = 'example.xlsx')

# Get the sheet names
print(wb.sheetnames)

# Get a sheet by name
sheet = wb['Sheet1']

# Print the sheet title
print(sheet.title)

# Get the currently active sheet or None
print(wb.active)

# Change active worksheet
wb.active = 2
print(wb.active)
```
```python
# ex1.0.3 Retrieving Cell Values

from openpyxl import load_workbook

wb = load_workbook(filename = 'example.xlsx')
sheet = wb['Sheet1']

# Select element A1 of the active sheet
print(sheet['A1'].value)

# Return the actual value of a cell
print(sheet['A1'].value)

# Return Value, Row and Column of element
rc = sheet['A1']
print(rc.value)
print('Row:', rc.row)
print('Column:', rc.column)
print(rc.coordinate)

# Return Value using cell
print(sheet.cell(row=1, column=1).value)
```
```python
# ex1.0.4 Retrieving Multiple Values
from openpyxl import load_workbook

wb = load_workbook(filename = 'example.xlsx')
sheet = wb['Sheet1']

# Get all cells from column A
print(sheet["A"])

# Get all cells for a range of columns
print(sheet["A:C"])

# Get all cells from row 1
print(sheet[1])

# Get all cells from row 1 to 3
print(sheet[1:3])

# Get all cells from row A1 to C3
print(sheet["A1:C3"])


# Select data with iter_rows
for row in sheet.iter_rows(min_row=1,max_row=4,min_col=1,max_col=3):
print(row)

# Get all whole dataset
for row in sheet.rows:
  print(row)

# Select data with iter_cols
for col in sheet.iter_cols(min_row=1,max_row=4,min_col=1,max_col=3):
  print(col)

# Get all whole dataset
for col in sheet.columns:
  print(col)

# Show values only
for col in sheet.iter_cols(min_row=1,max_row=4,min_col=1,max_col=3, values_only=True):
  print(col)
````
```python
# ex1.0.5 Converting data into Python structures
import json
from openpyxl import load_workbook

wb = load_workbook(filename = 'example.xlsx')

sheet = wb['Sheet1']

books = {}

for row in sheet.iter_rows(min_row=2,min_col=1,max_col=3,values_only=True):
    book_id = row[0]
    book = {
        "title": row[1],
        "author": row[2],
    }

    books[book_id] = book

print(json.dumps(books, indent=2))

```




https://www.datacamp.com/community/tutorials/python-excel-tutorial.0

https://realpython.com/openpyxl-excel-spreadsheets-python/