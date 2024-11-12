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

	' Loads an exl file 
	public sub loadFromFile(byval filePath, byval table, byval headers)
		Dim arrayDim
		arrayDim = UBound(headers)
		Dim temp
		temp = 0
		Dim element
		Dim SQL
		SQL = "SELECT "
		For Each element In headers
			If temp < arrayDim Then
        		SQL = SQl + "["&element&"], "
				Else 
    			SQL = SQl + "["&element&"] "
			End If
			temp = temp + 1
		Next
		Set ExcelConnection = Server.createobject("ADODB.Connection")
		ExcelConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.Open SQL, ExcelConnection
		'Response.Write "<table border=""1""><thead><tr>"
		temp = 0
		For Each Column In RS.Fields
		'Response.Write "<th>" & Column.Name & "</th>"
		'---xls.setHeader temp, Column.Name
		temp = temp + 1
		Next

		'Response.Write "</tr></thead><tbody>"
		IF NOT RS.EOF THEN
			temp = 0
			Dim temp2
			temp2 = 0 
			WHILE NOT RS.eof
				'Response.Write "<tr>"
				FOR EACH Field IN RS.Fields
					'Response.Write "<td>" & Field.value & "</td>"
					'---xls.setValue temp, temp2, Field.value
				NEXT
				'Response.Write "</tr>"
				RS.movenext
				temp = temp + 1
			WEND
	END IF
'Response.Write "</tbody></table>"
RS.close
ExcelConnection.Close
	end sub
end class
%>
