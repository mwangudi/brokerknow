<%
	Const report_ViewName = "OutstandingOrderList"
	Const reportPage = "OutstandingOrderList.asp"
	Const headerColCount = 3
	Const groupingHeaderCol = 0
	Const reportTitle = "Outstanding Orders"
%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%=reportTitle%></title>
	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	 <SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	
	<style media="print">
		@page {
			@top{font-family: Helvetica, Arial, sans-serif;
				font-size: 150%;
				font-weight: bolder;
				text-align: left;
				content: "<%= FormatDate(Date) %>";			
			}
			
			margin-left: 2cm;
			margin-right: 5cm;
			margin-top: 1cm;    
			margin-bottom: 2cm;
			size: portrait;
			
			br.newpage{
				page-break-before:always;
			}
			
			
		}

	</style>
</head>

<body Class="Reports">
<!--#include file="../libroutines.asp"-->

<%
Dim i
Dim headerColNames

genReport = Request.Form("genReport")
report_description = Request.Form("report_description")
selColumns = Request.Form("customCols")
useCustomization = Request.Form("useOwnFields")
SelectedSearchArgs = Request.Form("SelectedSearchArgs")
orderByCols = Request.Form("SelectedSortArgs")
headerColNames = Request.Form("headerColNames")

If genReport = "" Or report_description = "" Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm){
			
			frm.report_description.value = frm.genReport.options[frm.genReport.selectedIndex].text;
			report_validateFrm(frm);
		}		
	</Script>
	
	<form method="POST" action="<%=reportPage%>" Name="frmMain" id="frmMain">
		
		<input type="hidden" name="report_description" id="report_description">
		<input type="hidden" name="SelectedSearchArgs" id="SelectedSearchArgs">
		<input type="hidden" name="SelectedSortArgs" id="SelectedSortArgs">		
		<table>
			<tr>
				<td>Select type of report</td>
				<td>
					<SELECT Name="genReport">
						<OPTION VALUE="1">All</OPTION>
							
					</SELECT>
				</td>
			</tr>
			<tr>
				<td colspan="2"> <input type="checkbox" OnClick="JavaScript: switchDisplay (document.all.item('OwnFieldsSelectRow')); " class="BorderLess" name="useOwnFields" id="useOwnFields" value="1"> &nbsp; &nbsp; <label for="useOwnFields" style="cursor: hand">Customize this report (optional)</label></td>
			</tr>
			
			

			<tr style="display: none;" id="OwnFieldsSelectRow" align="right">
				<td colspan=2 valign="top" style="PADDING-LEFT: 25px">
					<TABLE border="0" cellpadding="2" cellspacing="5">
						<TR>
							<TD nowrap valign="top" STYLE="border: 1px ridge #000080">
 								Select columns to be included in the report
						<%
						Set Conn = GetActiveConnection("KBroker")
						Set Rs = Conn.OpenSchema (adSchemaColumns, Array(Empty, Empty, report_ViewName))
						
						
						If Not (Rs.EOF Or Rs.BOF) Then
							i = 1
							Do Until Rs.EOF
								if i > headerColCount then%>
										<br>
										<input type="checkbox" OnClick="JavaScript: evalCheck(this)" class="BorderLess" name="selColumns" value="[<%= Rs.Fields("COLUMN_NAME").Value %>]" id="<%= Rs.Fields("COLUMN_NAME").Value %>">
										<label for="<%= Rs.Fields("COLUMN_NAME").Value %>"><font face="Arial" color="navy" style="cursor: hand" TITLE="<%= Rs.Fields("COLUMN_NAME").Value %>"><%= Rs.Fields("COLUMN_NAME").Value %></font></label>
										
									<%
										
								else
										if trim(headerColNames) = "" then
												headerColNames = headerColNames & Rs.Fields("COLUMN_NAME").Value
										else
												headerColNames = "," & headerColNames & Rs.Fields("COLUMN_NAME").Value
										end if
										
										i = i + 1
								end if
								
								dataSearchValue = "[" & Rs.Fields("COLUMN_NAME").Value & "]:" & Rs.Fields("COLUMN_NAME").Value & "*0"
								dataSortValue = "[" & Rs.Fields("COLUMN_NAME").Value & "]:" & Rs.Fields("COLUMN_NAME").Value   
								If 	dataFldStr = "" Then
									dataFldStr =  dataSearchValue
									SortFieldsStr = dataSortValue   
								Else
									dataFldStr = dataFldStr & ";" & dataSearchValue 
									SortFieldsStr = SortFieldsStr & ";" & dataSortValue   
								End If
								
								Rs.MoveNext
							Loop
						End If%>
						<input type="hidden" name="headerColNames" id="headerColNames" value="<%=headerColNames%>">								
						</TD>
							<TD STYLE="border: 1px ridge #000080" valign="top">
								Change order of appearance of columns included in the report
								<TABLE border="0" cellpadding="2" cellspacing="5">
									<TR>
										<TD nowrap valign="top">
 
	
										<select name="customCols" multiple size="8">
															
											</select>
										</td>
										<td wrap valign="middle">	
											
										<img  STYLE="cursor: hand" class="buttons" TITLE="Move Up" src="../images/moveUp.gif" OnClick="JavaScript: moveUp();">
										<br>
										<img class="buttons" TITLE="Move Down" src="../images/moveDown.gif" OnClick="JavaScript: moveDown();" STYLE="cursor: hand">
																			
										</td>
									</tr>
								</table>		
					</TD>
					</tr>
					<tr>
						<td>
							<TABLE border="0" cellpadding="2" cellspacing="5" width="100%">
							<TR>
								<TD nowrap valign="top" STYLE="border: 1px ridge #000080">
 									Add custom filters
 									<br>
 									
 									
 									<table ID="mainTable">
										<TR>
											<TD><font face=Verdana size=1><b>Search by</b></font>			
											</TD>
											<TD COLSPAN=2><font face=Verdana size=1><b>Qualification</b></font>			
											</TD>
										</TR>
										<Script Language="VBScript">											
											report_init_Filter "<%= dataFldStr  %>"		
										</Script>	
									</table>
 				
 				
 						</td>
 							</tr>
 							</table>	
 						</td>
 						<td valign="top">
 							<TABLE border="0" cellpadding="2" cellspacing="5" width="100%">
							<TR>
								<TD nowrap valign="top" STYLE="border: 1px ridge #000080">
 									Add custom sort
 									<br>
 									
									 									
									<table border=0 cellspacing=1 cellpadding=2 width=100%>
										<TR>
											<TD COLSPAN=2>
												<table border=0 cellspacing=0 cellpadding=1>
													<tr>
														<td width=25%><font face=Verdana size=1>Sort by</font></td>
														<td width=100% align=left nowrap>&nbsp;</td>
													</tr>
												</table>	
											</TD>
											
										</TR>
										<Script Language="VBScript">											
											report_initSort "<%= SortFieldsStr %>"					
										</Script>
										
									</table>

								</td>
							</tr>
							</table>	
						</td>
					</tr>
					
					</table>
				</td>
			</tr>
			
			
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.getElementById('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>	
	
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>

<% DrawPageFunctions True, True, True 


If useCustomization <> "1" Then
	selColumns = "*"
	orderByCols = ""
	SelectedSearchArgs = ""
Else
	If selColumns = "" Then
		selColumns = "*"
	else
		selColumns = headerColNames & "," & selColumns		
	End If
	
	If orderByCols <> "" Then
		orderByCols = " ORDER BY " & headerColNames & "," & orderByCols 	
	End If	
	
	If SelectedSearchArgs  <> "" Then
		SelectedSearchArgs = "WHERE " & SelectedSearchArgs 
	End If
	
End If	



Select Case  genReport 
	Case "1"
		sqlStr = "SELECT " & selColumns & " FROM " & report_ViewName & " " & SelectedSearchArgs & " " &  orderByCols 
	Case "2"
		If SelectedSearchArgs <> ""  Then
			SelectedSearchArgs = " AND " & Mid(SelectedSearchArgs,  7)
		End If
		sqlStr = "SELECT " & selColumns & " FROM " & report_ViewName  & " WHERE [CDS Number] IS NOT NULL " & SelectedSearchArgs & " " &  orderByCols 
	Case "3"
		If SelectedSearchArgs <> ""  Then
			SelectedSearchArgs = " AND " & Mid(SelectedSearchArgs,  7)
		End If
		sqlStr = "SELECT " & selColumns & " FROM " & report_ViewName & " WHERE [CDS Number] IS NULL "  & SelectedSearchArgs & " " &  orderByCols  
	Case Else
		Response.End
End Select

	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")						        	
	'sqlStr = "SELECT * FROM ClientPortfolioStatement"
	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	'Rs.Filter = "Client_DPA_ LIKE '" & selectedClient & "' AND TransDate >= '" & FormatDate(selectedFromDate) & "'"
	
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("There are no transactions in the system")
			window.history.go(-1);
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
	
	
	Dim GroupingColumn
	GroupingColumn = Rs.Fields(groupingHeaderCol).Value  	
		
Do Until Rs.EOF 
%> 
<table border="0" cellspacing="2" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
		<td width="10%" nowrap><font face="Impact" size="4"><%=reportTitle%></font></td>
      <td width="60%" nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

  </table>

 
  
<br>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="1%"><b>Date:</b></td>
      <td width="48%"><%= FormatDate(Date) %></td>
    </tr>
	<%For i = 1 To headerColCount%>
			<tr>
			  <td ><b><%= Rs.Fields(i - 1).Name %>:</b></td>
			  <td ><b><%= Rs.Fields(i - 1).Value %></b></td>
			</tr>
	<%next%>
</table>
<BR>


  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="100%">
    <tr>
		<%For i = headerColCount To Rs.Fields.Count - 1
				if i = headerColCount then%>
					<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3"><%= Rs.Fields(i).Name %>:</font></b></td>
				<%elseif i = Rs.Fields.Count - 1 then%>
					<td style="border-right-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3"><%= Rs.Fields(i).Name %>:</font></b></td>			
				<%else%>
					<td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3"><%= Rs.Fields(i).Name %>:</font></b></td>
				<%end if%>
				
		<%next%>
      
    </tr>
    
<%
    nextEntityFound = False    
    
    Do Until nextEntityFound%>
		<tr>	
		  <%For i = headerColCount To Rs.Fields.Count - 1%>
					<td><%= Rs.Fields(i).Value %></td>
				
			<%next%>
		</tr>
	
	<%	Rs.MoveNext
		If Not Rs.EOF  Then
			If StrComp(GroupingColumn, Rs.Fields(groupingHeaderCol).Value) <> 0 Then
				GroupingColumn = Rs.Fields(groupingHeaderCol).Value
				nextEntityFound = True				
			End If   
		Else
			nextEntityFound = True	
		End If
		
	Loop
	%>
	
		<tr>	
		  <%For i = headerColCount To Rs.Fields.Count - 1%>
					<td>&nbsp;</td>
				
			<%next%>
		</tr>
	
    <tr>
      <td colspan="<%=Rs.Fields.Count%>" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp; </td>

    </tr>

   

  </table>
  
  <%If Not Rs.EOF Then %>
			<BR class="newpage">
	<%	End If
   

Loop

Set Rs = Nothing
Set Conn = Nothing%>   
</body>

</html>
