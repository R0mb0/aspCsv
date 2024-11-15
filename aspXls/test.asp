<%
option explicit

%><!--#include file="aspCsv.class.asp" --><!DOCTYPE HTML>
<html lang="en-US">
<head>
	<meta charset="ISO-8859-1">
	<title>aspCsv Tests</title>
	<link href="//netdna.bootstrapcdn.com/twitter-bootstrap/2.2.1/css/bootstrap.min.css" rel="stylesheet">
	<style type="text/css">
		.row-spaced-top {
			margin-top: 20px;
		}
	</style>
</head>
<body>
	<div class="navbar navbar-inverse navbar-static-top">
      <div class="navbar-inner">
        <div class="container">
          <div class="brand">aspCsv &diamond; Usage Tests</div>
        </div>
      </div>
    </div>
	<div class="container">
		<div class="row">
			<div class="span12">
				<h1 class="page-header">Examples</h1>
				<%
				dim startTime, csv, i
				startTime = timer
				
				set csv = new aspCsv
				
				' Add a header
				csv.setHeader 0, "id"
				csv.setHeader 1, "description"
				csv.setHeader 2, "createdAt"
				
				' Add the first data row
				csv.setValue 0, 0, 1
				csv.setValue 1, 0, "obj 1"
				csv.setValue 2, 0, date()
				
				' Add a bigger imcomplete data row
				csv.setValue 3, 1, "Comment"
				
				' Add a range of values at once
				csv.setRange 0, 2, Array(2, "obj 2", #11/25/2012#)
				
				
				' Add lots of data
				for i = 3 to 1000
					csv.setValue 0, i, i
					csv.setValue 1, i, "obj " & i
					csv.setValue 2, i, now()
				next
				
				
				dim outputCSV, outputTSV, outputHTML, outputPrettyHTML
				
				outputCSV = csv.toCSV()
				outputTSV = csv.toTabSeparated()
				
				outputHTML = csv.toHtmlTable()
				
				csv.prettyPrintHTML = true
				outputPrettyHTML = csv.toHtmlTable()
				
				%>
				<h2><code>toCSV()</code></h2>
				<pre class="pre-scrollable"><%= outputCSV %></pre>
				
				<h2><code>toTabSeparated()</code></h2>
				<pre class="pre-scrollable"><%= outputTSV %></pre>
				
				<h2><code>toHtmlTalbe() - prettyPrintHTML = false (default)</code></h2>
				<pre class="pre-scrollable"><%= Server.HtmlEncode(outputHTML) %></pre>
				
				<h2><code>toHtmlTalbe() - prettyPrintHTML = true</code></h2>
				<pre class="pre-scrollable"><%= Server.HtmlEncode(outputPrettyHTML) %></pre>
				
				<h4>HTML</h4>
				<%= replace(replace(replace(outputHTML, "<table", "<table class=""table table-striped table-bordered"""), "<thead><tr>", "<thead><tr class=""alert-info"">"), "<tbody", "<tbody class=""pre-scrollable""") %>

				<h4>Load informations from file</h4>
				<%
				csv.loadFromFile("file_example_csv_10.csv")
				%>

                		<h4>Table loaded:</h4>
				<%
				response.write(csv.toHtmlTable())
				%>

				<h4>Retrieve column's values from "Last Name"</h4>
				<ul>
				<%
                		dim temp
				for each temp in csv.getColumnValues("Last Name")
				response.write("<li>" & temp & "</li>")
				next
				%>
				</ul>

				<h4>Retreive rows's values from "3"</h4>
				<ul>
				<%
				for each temp in csv.getrowValues(3)
				response.write("<li>" & temp & "</li>")
				next
				%>
				</ul>

				<h4>Retreive a cell value from "Country3"</h4>
				<ul>
				<%
				response.write("<li> " & csv.getCellValue("Country", 3) & "</li>")
				%>
				</ul>
				
				<%	
				set csv = nothing
				%>
			</div>
		</div>
	</div>
	<div class="navbar navbar-inverse navbar-static-bottom">
		<div class="navbar-inner">
			<div class="container">
				<p></p>
				<p class="text-info pull-right">
					<i class="icon-time icon-white"></i> Generated in <%= clng((timer - startTime) * 1000) %>ms
				</p>
			</div>
		</div>
	</div>
</body>
</html>
