# aspCSV

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/d742480f427e475899203d858580bb7b)](https://app.codacy.com/gh/R0mb0/aspCsv/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/aspCsv)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/aspCsv)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

A classic ASP class that supports building, read and writing CSV, TSV (tab separated values)
and HTML outputing, including a pretty print for HTML.

> From: test.asp

## Easy to use

### Instantiate the class:

```asp
set csv = new aspCsv
```

### Create the table

**Add the header/titles for your structure:**
	
```asp
' Add a header: setHeader(x, value)
    csv.setHeader 0, "id"
    csv.setHeader 1, "description"
    csv.setHeader 2, "createdAt"
```
    
**Add some data:**

```asp
' Add the first data row: setValue(x, y, value)
    csv.setValue 0, 0, 1
    csv.setValue 1, 0, "obj 1"
    csv.setValue 2, 0, date()
    
' Add a range of values at once: setRange(x, y, valuesArray)
    csv.setRange 0, 2, Array(2, "obj 2", #11/25/2012#)

```

### Easy for reading:

> This will override the old informations

```asp
'Load a file:loadFromFile(path)
   csv.loadFromFile("file_example_csv_10.csv")
```

### Access informations 

**Check if a column exist:**

```asp
Check column existence: checkHeader(header)
   If csv.checkHeader("description") Then Response.write("everything is okay ")
```

**Get all values of a column:**

```asp
'Get column's values: getColumnValues(header)
   Dim cValues
   cValues = csv.getColumnValues("description")
```

**Get all values of a row:**

```asp
'Get row's values: getRowValues(row)
   Dim rValues
   rValues = csv.getRowValues(0)

```

**Get a cell's value:**

```asp
'Get a cell's value: getCellValue(header, row)
   Dim value
   value = csv.getCellValue "description", 0
```
 
### Easy for outputing:
	
**Output the data in string formatted values:**
	
```asp
outputCSV = csv.toCSV()
outputTSV = csv.toTabSeparated()
    
outputHTML = csv.toHtmlTable()
    
csv.prettyPrintHTML = true
outputPrettyHTML = csv.toHtmlTable()
```

**Or write it directly to a file:**

```asp
' Write the output to a file: writeToFile(filePath, format)
	csv.writeToFile("c:\mydata.csv", ASPcsv_CSV)
```
	
**The format flags supported are:**
	
```asp
ASPcsv_CSV = 1	' CSV format
ASPcsv_TSV = 2	' Tab separeted format
ASPcsv_HTML = 3	' HTML table format
```

