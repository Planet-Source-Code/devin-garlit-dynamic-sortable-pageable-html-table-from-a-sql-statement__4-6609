<div align="center">

## Dynamic, Sortable, Pageable, HTML table from a SQL statement


</div>

### Description

This is code for a dynamically created, sortable, pageable HTML table, this is

a pretty stripped down version. Real simple, just call the procedure where you

want the table, pass it a connection object and a SQL string, it will create an

ADO recordset and fill it into an HTML table, it will be fully pageable and sortable

by clicking the column head. You can also have the values in one column linkable to

another page,(example being you have an offer number and you click it to go to a details page) You input the records per page, default sort order, and the HTML tables attributes.

This can be easily made to incorporate images for column heads and for navigation buttons(maybe i'll post that later if this get a good response) Please email me with any questions.
 
### More Info
 
objConn = a connection object

strSQL = a string of SQL

strDefaultSort = a string of the default sorting column (i.e "FirstName")

intPageSize = integer of the number of records per page

strLinkedColumnName = a string of the colum to place a link on

strLink = a string of the page link

strTableAttributes = a string of HTML table attributes i.e. "name=myTable bgColor=steeleblue"

This is not that visually appealing, just shows the basics.

writes a sortable, pagable html table fill with records from a query


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Devin Garlit](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/devin-garlit.md)
**Level**          |Beginner
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Controls/ Forms/ Dialogs/ Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/controls-forms-dialogs-menus__4-3.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/devin-garlit-dynamic-sortable-pageable-html-table-from-a-sql-statement__4-6609/archive/master.zip)





### Source Code

'**************************************************************
'Function: createSortableList(objConn,strSQL, strDefaultSort, intPageSize, strLinkedColumnName,strLink,strTableAttributes)
'
'Returns: writes a sortable, pagable html table fill with records from a query
'
'Inputs:
'			objConn = a connection object
' strSQL = a string of SQL
' strDefaultSort = a string of the default sorting column (i.e "FirstName")
' intPageSize = integer of the number of records per page
' strLinkedColumnName = a string of the colum to place a link on
' strLink = a string of the page link
' strTableAttributes = a string of HTML table attributes i.e. "name=myTable bgColor=steeleblue"
'
'Sample Call:
'			createSortableList objConn,strSQL,"EmployeeID",3,"EmployeeID","employee_detail.asp","border=1 bgcolor='#cccccc'"
'
'Notes:
'
'			This is code for a dynamically created, sortable, pageable HTML table, this is
' a pretty stripped down version. Real simple, just call the procedure where you
' want the table, pass it a connection object and a SQL string, it will create an
' ADO recordset and fill it into an HTML table, it will be fully pageable and sortable
' by clicking the column head. You can also have the values in one column linkable to
'			another page,(example being you have an offer number and you click it to go to a details page)
' You input the records per page, default sort order, and the HTML tables attributes.
' This can be easily made to incorporate images for column heads and for navigation buttons(maybe i'll
' post that later if this get a good response) Please email me with any questions.
'
'Programmer: Devin Garlit (dgarlit@hotmail.com) 4/25/01
'**************************************************************
sub createSortableList(objConn,strSQL, strDefaultSort, intPageSize, strLinkedColumnName,strLink,strTableAttributes)
		dim RS,strSort, intCurrentPage, strPageName
		dim strTemp, field, strMoveFirst, strMoveNext, strMovePrevious, strMoveLast
		dim i, intTotalPages, intCurrentRecord, intTotalRecords
		i = 0
		strSort = request("sort")
		intCurrentPage = request("page")
		strPageName = Request.serverVariables("SCRIPT_NAME")
		if strSort = "" then
			strSort = strDefaultSort
		end if
		if intCurrentPage = "" then
			intCurrentPage = 1
		end if
		set RS = server.CreateObject("adodb.recordset")
		with RS
			.CursorLocation=3
			.Open strSQL & " order by " & replace(strSort,"desc"," desc"), objConn,3 '3 is adOpenStatic
			.PageSize = cint(intPageSize)
			intTotalPages = .PageCount
			intCurrentRecord = .AbsolutePosition
			.AbsolutePage = intCurrentPage
			intTotalRecords = .RecordCount
		end with
		Response.Write "<table " & strTableAttributes & " >" & vbcrlf
		'table head
		Response.Write "<tr>" & vbcrlf
		for each field in RS.Fields 'loop through the fields in the recordset
			Response.Write "<td align=center>" & vbcrlf
			if instr(strSort, "desc") then 'check the sort order, if its currently ascending, make the link descending
				Response.Write "<a href=" & strPageName & "?sort="& field.name & "&page=" & intCurrentPage & ">" & field.name & "</a>" & vbcrlf
			else
				Response.Write "<a href=" & strPageName & "?sort="& field.name &"desc&page=" & intCurrentPage & ">" & field.name & "</a>"	& vbcrlf
			end if
			Response.Write "<td>"	& vbcrlf
		next
		Response.Write "<tr>"
		'records
		for i = intCurrentRecord to RS.PageSize 'display from the current record to the pagesize
			if not RS.eof then
			Response.Write "<tr>" & vbcrlf
			for each field in RS.Fields 'for each field in the recordset
				Response.Write "<td align=center>" & vbcrlf
				if lcase(strLinkedColumnName) = lcase(field.name) then 'if this field is the "linked field" provide a link
					Response.Write "<a href=" & strLink & "?sort="& strSort &"&page=" & intCurrentPage & "&" & field.name & "=" & field.value & " >" & field.value & "</a>" & vbcrlf
				else
					Response.Write field.value
				end if
				Response.Write "<td>" & vbcrlf
			next
			Response.Write "<tr>" & vbcrlf
			RS.MoveNext
			end if
		next
		Response.Write "<table>" & vbcrlf
		'page navigation
		select case cint(intCurrentPage)
			case cint(intTotalPages) 'if its the last page give only links to movefirst and move previous
				strMoveFirst = "<a href=" & strPageName & "?sort="& strSort &"&page=1 >"& "First" &"</a>"
				strMoveNext = ""
				strMovePrevious = "<a href=" & strPageName & "?sort="& strSort &"&page=" & intCurrentPage - 1 & " >"& "Prev" &"</a>"
				strMoveLast = ""
			case 1 'if its the first page only give links to move next and move last
				strMoveFirst = ""
				strMoveNext = "<a href=" & strPageName & "?sort="& strSort &"&page=" & intCurrentPage + 1 & " >"& "Next" &"</a>"
				strMovePrevious = ""
				strMoveLast = "<a href=" & strPageName & "?sort="& strSort &"&page=" & intTotalPages & " >"& "Last" &"</a>"
			case else
				strMoveFirst = "<a href=" & strPageName & "?sort="& strSort &"&page=1 >"& "First" &"</a>"
				strMoveNext = "<a href=" & strPageName & "?sort="& strSort &"&page=" & intCurrentPage + 1 & " >"& "Next" &"</a>"
				strMovePrevious = "<a href=" & strPageName & "?sort="& strSort &"&page=" & intCurrentPage - 1 & " >"& "Prev" &"</a>"
				strMoveLast = "<a href=" & strPageName & "?sort="& strSort &"&page=" & intTotalPages & " >"& "Last" &"</a>"
		end select
		with Response
			.Write strMoveFirst & " "
			.Write strMovePrevious
			.Write " " & intCurrentPage & " of " & intTotalPages & " "
			.Write strMoveNext & " "
			.Write strMoveLast
		end with
		if RS.State = &H00000001 then 'its open
			RS.Close
		end if
		set RS = nothing
	end sub

