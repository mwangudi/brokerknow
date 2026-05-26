<html>
	<head>
		<meta http-equiv="Content-Language" content="en-uk">
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
		<title>View Report</title>
				 
				 
		</head>

		<body Class="Dialog">

<!--#include file="../libroutines.asp"-->


<%
	Dim ID, userID
	const UDLName = "KBroker"
   
	ID = Request("ID")
	userID = Session("UserID")
	
	If (Trim(ID) = "") or (ID = "0") Then%>
			<script language = 'JavaScript'>
                	alert ("Please select a Report");
                	window.parent.self.close();
			</script>
			<%Response.End
	End If
	
	If userID = "" Then
		Response.End
	End If
	
	sqlStr = "SELECT * FROM Menus WHERE MenuID = " & ID & " AND IsReport = 1 AND EXISTS(SELECT     MenuGroups.ID " & _
				" FROM         UserGroups INNER JOIN " & _
				"                      MenuGroups ON UserGroups.GroupID = MenuGroups.groupID " & _
				"			WHERE     (UserGroups.UserID = " & userID & ") AND (MenuGroups.MenuID = Menus.menuID))  ORDER BY  mnuCaption"
		
	Set conn = GetActiveConnection(UDLName)
		        
    Set rs = conn.Execute(sqlStr)
    
    If Not (Rs.EOF Or Rs.BOF) Then
		Response.Redirect "../" & Rs.Fields("mnuAction").Value		
	Else%>
			<Script Language="JavaScript">
				alert("The specified report cannot be accessed. Ensure that you have permission to view this report")
				window.self.close();
			</Script>
		<%
	End If	
	
	Set Rs = Nothing
	Set Conn = Nothing
%>



</body>

</html>
