<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Reports</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	 <SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	 <SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
</head>
<body Class="Reports">

<!--#include file="../libroutines.asp"-->

<%

reportsURL = Request.QueryString("rURL") 
mnuID = Request.QueryString("id")

If mnuID <> "" And reportsURL <> "" Then
	Set Conn = GetActiveConnection("KBroker")
	SQL = "SELECT * FROM Menus WHERE ReportsMenuID = " & mnuID
	Set Rs = Conn.Execute(SQL)
	If Not (Rs.EOF Or Rs.BOF) Then
		ReportObjName = Rs.Fields("mnuReportObjectName").Value		
	End If
	
	If ReportObjName = "" Or IsNull(ReportObjName) Then
		Response.Redirect Unescape(reportsURL) 
		Response.End 
	Else
	
	
%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm){									
			frm.report_description.value = frm.genReport.options[frm.genReport.selectedIndex].text;
			report_validateFrm(frm);
		}
		
		
	</Script>
	<form method="POST" action="<%= reportsURL %>" Name="frmMain" id="frmMain">
		<input type="hidden" name="report_description" id="report_description">
		<input type="hidden" name="SelectedSearchArgs" id="SelectedSearchArgs">
		<input type="hidden" name="SelectedSortArgs" id="SelectedSortArgs">	
		<table>
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
						SQL = "SELECT * FROM _ReportsParameters_ WHERE ReportsMenuID = " & mnuID
						Set Rs = Conn.Execute(SQL)
						If Not (Rs.EOF Or Rs.BOF) Then
							Do Until Rs.EOF%>
								<br>
								<input type="checkbox" OnClick="JavaScript: evalCheck(this)" class="BorderLess" name="selColumns" value="[<%= Rs.Fields("COLUMN_NAME").Value %>]" id="<%= Rs.Fields("COLUMN_NAME").Value %>">
								<label for="<%= Rs.Fields("COLUMN_NAME").Value %>"><font face="Arial" color="navy" style="cursor: hand" TITLE="<%= Rs.Fields("COLUMN_NAME").Value %>"><%= Rs.Fields("COLUMN_NAME").Value %></font></label>
							<%	
							
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
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.getElementById('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%	Set rs = Nothing
		Set Conn = Nothing
		Response.End
	End If   
End If

</body>

</html>
