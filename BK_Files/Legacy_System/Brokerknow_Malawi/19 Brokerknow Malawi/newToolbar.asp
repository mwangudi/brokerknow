<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="STYLE/webparts.css"> 
<title>Knowing Toolbar</title>
<base target="rbottom">
</head>

<body background="images/menuGradient.JPG">


<!--#include file="libroutines.asp"-->
<%

Dim Conn

Dim RS
Dim m_mnuArray(100)
Dim mnuCaption
Dim rsChild, rsSubChild 

UserID = Session("UserID")

If UserID = "" Then
	'session expired
	Response.Write "Session Expired"
	Response.End 
End If


Set Conn = GetActiveConnection("KBroker")
mnuID = Request.QueryString("mnuID")

 SQL = "SELECT * FROM Menus WHERE Abs(isMainMenu) <> 1 AND mainmenuID = '" & mnuID & "' AND EXISTS(SELECT     MenuGroups.ID " & _
			" FROM         UserGroups INNER JOIN " & _
			"                      MenuGroups ON UserGroups.GroupID = MenuGroups.groupID " & _
			"			WHERE     (UserGroups.UserID = " & userID & ") AND (MenuGroups.MenuID = Menus.menuID))   ORDER BY mnuCaption"

Set Rs = Conn.Execute(SQL) %>

<div align="right">

<table  border="0" style="font-family: Tahoma; font-size: 8pt;" cellpadding="2" cellspacing="0">
  <TR>
	<TD  align="left" nowrap width="2%">
		&nbsp;
										
			<IMG SRC="images/back.jpg" STYLE="cursor: hand" title="Back" OnClick="JavaScript: window.parent.frames[2].history.go(-1);">
		
		&nbsp;
		<IMG SRC="images/forward.jpg" STYLE="cursor: hand" title="Forward" OnClick="JavaScript: window.parent.frames[2].history.go(1);">
	</TD>
  </TR>
</table>
<table  border="0" style="font-family: Tahoma; font-size: 8pt;" cellpadding="2" cellspacing="0">
  <tr>
  	<TD  valign="top" nowrap>
	<%If Not (Rs.EOF Or Rs.BOF) Then 
  		Do Until Rs.EOF	
  		
  		
  			
              If Not (Rs.Fields("mnuaction").Value  = "") Then
                  Action = Rs.Fields("mnuaction").Value 
              Else
                  Action = "#"
              End If                
                
              If IsNull(Rs("Image")) Then
					mnuImage = "images/linksubnote.gif"
				ElseIf Rs("Image") = "" Then
					mnuImage = "images/linksubnote.gif"
				Else
					mnuImage = Rs("Image")
				End If		
				
				If InStr(1, Rs.Fields("mnucaption").value, "Add") > 0  Then
					myAddLink = Action 
					myAddImage = mnuImage 
					myAddID = Rs("menuID")
					myAddCaption =  Rs.Fields("mnucaption").value
				ElseIf InStr(1, Rs.Fields("mnucaption").value, "Delete") > 0  Then
					myDelLink = Action 
					myDelImage = mnuImage 
					myDelID = Rs("menuID")				
					myDelCaption =  Rs.Fields("mnucaption").value
				ElseIf InStr(1, Rs.Fields("mnucaption").value, "Edit") > 0  Then
					myEditLink = Action 
					myEditImage = mnuImage 
					myEditID = Rs("menuID")				
					myEditCaption =  Rs.Fields("mnucaption").value
				Else%>	
				
							
								&nbsp;
								<A href="<%= Action %>">								
										<%= Rs.Fields("mnucaption").value %>
								</A>
								&nbsp;
					
				
		<%		End If
		
			Rs.MoveNext
			
		Loop
		
	  End If
	
	Set Rs = Nothing
	Set Conn = Nothing%>  	
	</td>		
  </tr>
  </table>
  
  <table  border="0" valign=top style="font-family: Tahoma; font-size: 8pt;" cellpadding="2" cellspacing="0">
  <tr> 		
	<TD align="center" width="10%" nowrap>								
		&nbsp;
		<A href="Search.asp">
			Search
		</A>
		&nbsp;
		<A href="Help.asp">
			Help
		</A>
		&nbsp;
	</TD>
	<TD align="right" width="90%" nowrap>
  		<%For i = 1 To 3
  			Select Case i
  				Case 1
  					'add
  					Action = myAddLink 
					mnuImage = myAddImage  
					mnuID = myAddID 
					'mnuCaption = myAddCaption 
					mnuCaption = "Add"
				Case 2
  					'delete
  					Action = myDelLink 
					mnuImage = myDelImage 
					mnuID = myDelID 
					'mnuCaption = myDelCaption 
					mnuCaption = "Delete"
				Case 3
  					'edit
  					Action = myEditLink 
					mnuImage = myEditImage 
					mnuID = myEditID 
					'mnuCaption = myEditCaption 
					mnuCaption = "Edit"					
			End Select	 
			
			If Action <> "" Then %>
								&nbsp;							
					  			<A href="<%= Action %>">
									<%= mnuCaption  %>
								</A>
								&nbsp;
		
			<%
			Else
				Response.Write "&nbsp;"
			End If
		Next%>
		</TD>
  </tr>
</table>



</div>



<SCRIPT LANGUAGE="JavaScript">
	function Change(row){
		document.all.item("IMG" + row.name).filters.alpha.opacity=100;	
	}
	
	function UnChange(row){
		document.all.item("IMG" + row.name).filters.alpha.opacity=50;
	}
	
</Script>


</body>

</html>
