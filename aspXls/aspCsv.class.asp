<%
' Classic ASP CSV creator 1.0
' By RCDMK <rcdmk@rcdmk.com>
'
' The MIT License (MIT) - http://opensource.org/licenses/MIT
' Copyright (c) 2012 RCDMK <rcdmk@rcdmk.com>
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
' associated documentation files (the "Software"), to deal in the Software without restriction,
' including without limitation the rights to use, copy, modify, merge, publish, distribute,
' sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial
' portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
' NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
' NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
' DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT
' OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


' Format constants
const ASPXLS_CSV = 1	' CSV format
const ASPXLS_TSV = 2	' Tab separeted format
const ASPXLS_HTML = 3	' HTML table format


' Main class
class aspExl
	dim lines(), curBoundX, curBoundY
	dim headers()
	dim m_prettyPrintHTML
	dim headerIndex
	
	
	' A flag for outputing more readable HTML
	public property get prettyPrintHTML()
		prettyPrintHTML = m_prettyPrintHTML
	end property
	
	public property let prettyPrintHTML(byval value)
		m_prettyPrintHTML = value
	end property
	
	
	' Initialization and destruction
	sub class_initialize()
		curBoundX = -1
		curBoundY = -1
		m_prettyPrintHTML = false
	end sub
	
	sub class_terminate()
		' Destroy all elements of the arrays
		redim lines(-1)
		redim headers(-1)
	end sub
	
	
	' Resizes the columns and header arrays to fit a new size
	private sub resizeCols(byval newSize)
		dim i, cols
		
		for i = 0 to curBoundY
			cols = lines(i)
			redim preserve cols(newSize)
			lines(i) = cols
		next
		
		redim preserve headers(newSize)
		
		curBoundX = newSize
	end sub
	
	
	' Resizes the lines array to fit a new size
	private sub resizeRows(byval newSize)
		dim i
		redim preserve lines(newSize)
		
		for i = curBoundY + 1 to newSize
			if i >= 0 then lines(i) = Array()
		next
		
		curBoundY = newSize
		
		resizeCols curBoundX
	end sub
	
	
	' Contatenates and return the values in a string using a separator string
	private function toString(byval separator)
		dim output, headersString, i
		output = ""
		headersString = join(headers, separator)
		
		if replace(headersString, separator, "") <> "" then output = headersString & vbCrLf
		
		for i = 0 to curBoundY
			output = output & join(lines(i), separator) & vbCrLf
		next
		
		toString = output
	end function
	

	' Sets a header value
	public sub setHeader(byval x, byval value)
		if x > curBoundX then resizeCols(x)
		
		headers(x) = value
	end sub

	' Sets the value of a cell
	public sub setValue(byval x, byval y, byval value)
		dim cols
		
		if y > curBoundY then resizeRows y		
		if x > curBoundX then resizeCols x
		
		cols = lines(y)
		cols(x) = value
		
		lines(y) = cols
	end sub
	
	
	' Sets the values of a range of cells starting at the specified coordinates
	public sub setRange(byval x, byval y, byval arr)
		if y > curBoundY then resizeRows y
		
		dim arrBound
		arrBound = ubound(arr)
		
		if arrBound + x > curBoundX then resizeCols(arrBound + x)
		
		dim i, cols
		cols = lines(y)
		
		for i = 0 to arrBound
			cols(x + i) = arr(i)
		next
		
		lines(y) = cols
	end sub
	
	
	' Returns a string formatted output of the data
	public function outputTo(byval format)
		dim output, headersString, i
		
		select case format
			case ASPXLS_HTML:
				output = "<table>"
				headersString = join(headers, "</th><th>")
				
				if replace(headersString, "</th><th>", "") <> "" then output = output & "<thead><tr><th>" & headersString & "</th></tr></thead>"
				
				output = output & "<tbody>"
				
				for i = 0 to curBoundY
					output = output & "<tr><td>" & join(lines(i), "</td><td>") & "</td></tr>"
				next
				
				output = output & "</tbody></table>"
				
				
				' Prettify HTML for easy reading
				if m_prettyPrintHTML then
					dim lineSeparator, indentChar
					dim regex, breakAndIndent, doubleIndent

					lineSeparator = vbCrLf
					indentChar = vbTab
					
					breakAndIndent = lineSeparator & indentChar
					doubleIndent = indentChar & indentChar
					
					set regex = new regexp
					regex.global = true
					
					regex.pattern = "(</?(?:thead|tbody)>)"
					output = regex.replace(output, breakAndIndent & "$1")
					
					regex.pattern = "(</?tr>)"
					output = regex.replace(output, breakAndIndent & indentChar & "$1")
					
					regex.pattern = "(</table>)"
					output = regex.replace(output, lineSeparator & "$1")
					
					regex.pattern = ">(<(?:th|td)>)"
					output = regex.replace(output, ">" & breakAndIndent & doubleIndent & "$1")
					
					set regex = nothing
				end if
				
			case ASPXLS_CSV:
				output = toString(";")
				
			case ASPXLS_TSV:
				output = toString(vbTab)
			
			case default:
				output = ""
		end select
		
		outputTo = output
	end function
	
	
	' Returns a semi-colon separated string for each row
	public function toCSV()
		toCSV = outputTo(ASPXLS_CSV)
	end function
	
	
	' Returns a TAB separated string for each row
	public function toTabSeparated()
		toTabSeparated = outputTo(ASPXLS_TSV)
	end function
	
	
	' Returns a HTML table formatted string
	public function toHtmlTable()
		toHtmlTable = outputTo(ASPXLS_HTML)
	end function
	
	
	' Writes a string fomatted output to a file
	public sub writeToFile(byval filePath, byval format)
		dim fso, file
		dim i
		set fso = createObject("scripting.filesyStemObject")
		
		set file = fso.createTextFile(filePath, true)
		
		file.writeLine outputTo(format)
		
		file.close
		set file = nothing
		
		set fso = nothing
	end sub

	' Loads an csv file 
	public sub loadFromFile(byval filePath)
		'Reset the status
		class_initialize()
		'Start reading
		Set fso = Server.CreateObject("Scripting.FileSystemObject") 
		Set fs = fso.OpenTextFile(Server.MapPath("file_example_XLS_10.csv"), 1, true)
		Dim index
		index = 0
		Dim temp_array
		Dim temp
		Dim temp_index
		Do Until fs.AtEndOfStream
			temp_array = Split(fs.ReadLine, ",")
			If index = 0 Then 
				temp_index = 0
				For Each temp In temp_array
					setHeader temp_index, temp 
					temp_index = temp_index + 1
				Next
			Else
				temp_index = 0
				For Each temp In temp_array
					setValue temp_index, index, temp
					temp_index = temp_index + 1
				Next
			End If
			index = index + 1
		Loop
	end sub

	'Checks if a header is present and set the it's index
	public function checkHeader(byval header)
		Dim temp 
		Dim temp_index
		temp_index = 0
		For Each temp in headers
			If temp = header Then
				headerIndex = temp_index
				checkHeader = true
				Exit Function
			End If
			temp_index = temp_index + 1 
		Next
		checkHeader = false
	end function

	'Retreives the values of a column 
	public function getColumnValues(byval header)
		'Check if the header is present and get the index
		If Not checkHeader(header) Then
			Call Err.Raise(vbObjectError + 10, "aspCsv.class.asp - getColumnValues", "The header -" & header & "- does not exist")
		End If 
		'Create variables
        Dim temp_array
        temp_array = Array()
        Dim temp_array_index
        temp_array_index = 0
        Dim temp_line
		'extract values
        For Each temp_line In lines
            Redim Preserve temp_array(temp_array_index)
            temp_array(temp_array_index) = temp_line(headerIndex)
            temp_array_index = temp_array_index + 1
        Next
		'return
		getColumnValues = temp_array
	end function
	
	'Retreives the values of a row 
	public function getRowValues(byval row)
		If Not (row >=0 and row <= curBoundY) Then 
			Call Err.Raise(vbObjectError + 10, "aspCsv.class.asp - getRowValues", "The row -" & row & "- does not exist")
		End If
		getRowValues = lines(row)
	end function

	'Extract a cell value
	public function getCellValue(byval header, byval row)
		'Check if the header is present and get the index
		If Not checkHeader(header) Then
			Call Err.Raise(vbObjectError + 10, "aspCsv.class.asp - getCellValue", "The header -" & header & "- does not exist")
		End If 
		'Check if the row exist
		If Not (row >=0 and row <= curBoundY) Then 
			Call Err.Raise(vbObjectError + 10, "aspCsv.class.asp - getRowValues", "The row -" & row & "- does not exist")
		End If
		getCellValue = getColumnValues(header)(row)
	end function

end class
%>
