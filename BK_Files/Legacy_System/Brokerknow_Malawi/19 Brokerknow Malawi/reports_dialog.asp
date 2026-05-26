<html>

<head>
<title>Reports</title>
<LINK href="STYLE/default.css" type=TEXT/CSS rel=STYLESHEET> 
<LINK href="STYLE/webparts.css" type=TEXT/CSS rel=STYLESHEET>
<SCRIPT language=Javascript src="scripts/common.js"></SCRIPT>
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
</head>

<body class="Dialog">
<!--#include file="libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<%

mnuCatID = Request("ID")
UserID = Session("UserID")

If mnuCatID = "" Then%>
	<Script Language="JavaScript">
		alert("This process cannot continue: an unexpected error has occured.")
	</Script>
	<%
	
	WriteDialogRefuseOpenScript
	Response.End
End If



sqlStrOrig = "SELECT * FROM Menus WHERE MainMenuID = " & mnuCatID & " AND IsReport = 1 AND EXISTS(SELECT     MenuGroups.ID " & _
		" FROM         UserGroups INNER JOIN " & _
		"                      MenuGroups ON UserGroups.GroupID = MenuGroups.groupID " & _
		"			WHERE     (UserGroups.UserID = " & userID & ") AND (MenuGroups.MenuID = Menus.menuID))  ORDER BY  mnuCaption"
		
Set conn = GetActiveConnection("KBroker")	
Set Rs = Conn.Execute(sqlStrOrig) 

If Rs.EOF OR Rs.BOF Then%>
	<Script Language="JavaScript">
		alert("There are currently no reports configured under the selected operation")
	</Script>
	<%
	Set Rs = Nothing
	Set Conn = Nothing
	WriteDialogRefuseOpenScript
	Response.End
End If%>

<form method="POST" action="" Name="frmReports">
<input type="hidden" value="1" name="genReport">
<input type="hidden" Name="timeLimit" value="0" id="NSETime">


  <table width=100% cellpadding=3>
	<%If Not (Rs.EOF Or Rs.BOF) Then%>
	
		<tr>
			<td>
				<b>Date: </b>
				<SCRIPT language="JavaScript">			
					var cal=new ctlSpiffyCalendarBox("cal", "frmReports", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
					cal.writeControl();
				</SCRIPT>
			</td>
		</tr>	
		
		<%		
		mnuCount = 0
		Do Until Rs.EOF
			mnuCount = mnuCount  + 1%>
		<tr>
			<td><input type=radio class="BorderLess" value="<%= Rs.Fields("mnuAction").Value %>" NAME=selReport ID="OPT-<%= mnuCount %>"><LABEL FOR="OPT-<%= mnuCount %>" STYLE="CURSOR: HAND"><%= Rs.Fields("mnuCaption").Value %></LABEL>
			</td>
		</tr>
	<%		Rs.MoveNext
		Loop%>
		
		<tr>
			<td align="right" valign="bottom"> 
				<input type="button" class="Buttons" OnClick="VBScript: generateReport" value=" Generate ">&nbsp;&nbsp;&nbsp;<input type="button" value=" Close " OnClick="JavaScript: window.self.close();" class="Buttons">
			</td>
		</tr>	
	<%	
	End If%>	
  </table>	
		
  
</form>

<Script Language="VBScript">
	Function generateReport
		Dim frm, is_checked, targetPage
		Set frm = document.all.item("frmReports")
		is_checked = False
		For Each Thing In frm
			On Error Resume Next
			If Thing.Name = "selReport" Then
				If Thing.Checked = True Then
					is_checked = True
					targetPage = Thing.Value
					Exit For
				End If
			End If
		Next
		
		If is_checked = False Then
			alert "No report selected"
			Exit Function
		End If
		
		reportCount = reportCount + 1
		Set RepWin = window.open("", "report" & reportCount, "height=600,width=700,status=no,scrollbars=yes,toolbar=no,menubar=no,location=no,resizable=yes")
		RepWin.opener = window.self
		
		frm.action = targetPage
		frm.target = "report" & reportCount
		frm.submit 
		
		window.parent.self.close
	End Function
</Script>

<%
Set Rs = Nothing
Set Conn = Nothing%>

</body>

</html>
