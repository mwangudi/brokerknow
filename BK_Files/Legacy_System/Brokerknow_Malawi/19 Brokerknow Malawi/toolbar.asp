<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Knowing Toolbar</title>
<base target="maininfo">
</head>

<body background="images/menuGradient.JPG">


<!--#include file="libroutines.inc"-->
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

<table  border="1" align=top style="font-family: Tahoma; font-size: 8pt; color: #800000" cellpadding="2" cellspacing="0">
  <tr>
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
				End If		%>	
				
					<A href="<%= Action %>">	  						
							<TD name="<%= rs("menuID") %>" class="flyoutLink" handle="298" STYLE="cursor: HAND; background-image: url('images/loginback.JPG'); background-repeat: no-repeat; background-attachment: fixed; background-position: center 50%" bordercolor="#336699" bordercolorlight="#0066CC" bordercolordark="#FFFFFF" OnMouseOver="JavaScript: Change(this)" OnMouseOut="JavaScript: UnChange(this)">								
								<img src="<%= mnuImage %>" name="IMG<%= rs("menuID") %>" style="filter:alpha(opacity=50); -moz-opacity:0.5" width=20 height=20>&nbsp;<%= rS.Fields("mnucaption").value %>
							</TD>
					</A>

		<%
			Rs.MoveNext
		Loop
	  End If
	
	Set Rs = Nothing
	Set Conn = Nothing%>  
		
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
