# aspCSV

A classic ASP class that supports building, read and writing CSV, TSV (tab separated values)
and HTML outputing, including a pretty print for HTML.

> From: test.asp

## Easy to use

### Instantiate the class:

```
set xls = new aspExl
```

### Create the table

**Add the header/titles for your structure:**
	
```
' Add a header: setHeader(x, value)
    xls.setHeader 0, "id"
    xls.setHeader 1, "description"
    xls.setHeader 2, "createdAt"
```
    
**Add some data:**

```
' Add the first data row: setValue(x, y, value)
    xls.setValue 0, 0, 1
    xls.setValue 1, 0, "obj 1"
    xls.setValue 2, 0, date()
    
' Add a range of values at once: setRange(x, y, valuesArray)
    xls.setRange 0, 2, Array(2, "obj 2", #11/25/2012#)

```

### Easy for reading:

> This will override the old informations

```
'Load a file:loadFromFile(path)
   xls.loadFromFile("file_example_XLS_10.csv")
```

### Access informations 

**Check if a column exist**

```
Check column existence: checkHeader(header)
   If xls.checkHeader("description") Then Response.write("everything is okay ")
```

**Get all values of a column:**

```
'Get column's values: getColumnValues(header)
   Dim cValues
   cValues = xls.getColumnValues("description")
```

**Get all values of a row:**

```
'Get row's values: getRowValues(row)
   Dim rValues
   rValues = xls.getRowValues(0)

```

**Get a cell's value:**

```
'Get a cell's value: getCellValue(header, row)
   Dim value
   value = xls.getCellValue "description", 0
```
 
### Easy for outputing:
	
**Output the data in string formatted values:**
	
```
outputCSV = xls.toCSV()
outputTSV = xls.toTabSeparated()
    
outputHTML = xls.toHtmlTable()
    
xls.prettyPrintHTML = true
outputPrettyHTML = xls.toHtmlTable()
```

**Or write it directly to a file:**

```
' Write the output to a file: writeToFile(filePath, format)
	xls.writeToFile("c:\mydata.csv", ASPXLS_CSV)
```
	
**The format flags supported are:**
	
```
ASPXLS_CSV = 1	' CSV format
ASPXLS_TSV = 2	' Tab separeted format
ASPXLS_HTML = 3	' HTML table format
```

