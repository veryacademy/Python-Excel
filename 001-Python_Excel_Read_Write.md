# **001-Reading Excel files with Pandas read_excel()**

Welcome to the Python Excel Series.

**Video Tutorial:** https://youtu.be/bI68wnoINwc

#### **Difficulty:** Beginner
#### **Tags/Keywords:** Python, Excel, Pandas
---

### Code Examples

```python
# ex1.0.1 DataFrame Creation
import pandas as pd

excel_file = 'books.xlsx'
csv_file = 'books.csv'

df = pd.read_excel(excel_file)
dff = pd.read_csv(csv_file)
```
```python
# ex1.0.1 Excel Sheet to Dict, CSV and JSON
import pandas as pd

excel_file = 'books.xlsx'
df = pd.read_excel(excel_file)

new_dict = df.to_dict()
new_csv = df.to_csv(index=False)
new_json = df.to_json()

print(new_dict)
```
```python
# ex1.0.2 read_excel() parameter values
import pandas as pd

excel_file = 'books.xlsx'
df = pd.read_excel(excel_file)

# sheet_name='name'
# sheet_name=0
# usecols=['title', 'authors']

print(df.head(5))
```
```python
# ex1.0.2 Inspect DataFrame
import pandas as pd

excel_file = 'books.xlsx'
df = pd.read_excel(excel_file)

excel_file = 'books.xlsx'
df = pd.read_excel(excel_file)

print(df)
print(df.head(5))
print(df.tail(5))
print(df.index)
print(df.columns)
print(df.dtypes)
```
```python
# ex1.0.3 Selecting Data
import pandas as pd

df = pd.read_excel(excel_file)

excel_file = 'books.xlsx'
df = pd.read_excel(excel_file)

# df
# df.at can only access a single value at a time.
# df.loc can select multiple rows and/or columns.

# Series
# Series is the data structure for a single column of a DataFrame
print(df['title'])
print(df['title'].head(5))

# DataFrame
print(df[['title','authors']].head(5))
print(df.at[0, 'title']) 
print(df.loc[0:3, 'title':'authors'])
```
```python
# ex1.0.4 Select Sheets
import pandas as pd

excel_file = 'books.xlsx'
df = pd.read_excel(excel_file)

excel_file = 'books.xlsx'
df = pd.read_excel(excel_file, sheet_name=[0,1], nrows=5)

# sheet_name=0
# sheet_name="books_2"
# sheet_name=[0,1], nrows=5

# print(df)
# Specify a sheet to use
print(df[0].head(2))
```
```python
# ex1.0.5 Working with all sheets
import pandas as pd

excel_file = 'books.xlsx'
df = pd.read_excel(excel_file)

excel_file = 'books.xlsx'
df = pd.read_excel(excel_file, sheet_name=None)

print(df['books_2'].head(5))
```
---
### **References/Further Reading**
https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
https://pypi.org/project/xlrd/
https://pypi.org/project/openpyxl/
https://pythonbasics.org/read-excel/
https://stackoverflow.com/questions/26047209/what-is-the-difference-between-a-pandas-series-and-a-single-column-dataframe