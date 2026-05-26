<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Navigation Menus</TITLE>
<base target="maininfoR">

<LINK href="style/default.css" type=text/css rel=stylesheet>
<SCRIPT language=Javascript src="scripts/common.js"></SCRIPT>
<SCRIPT language=Javascript src="scripts/uplevel.js"></SCRIPT>
<LINK href="style/OSfxDocument.css" type=text/css rel=stylesheet>
<SCRIPT language=Javascript src="scripts/fhsupport.js"></SCRIPT>

<SCRIPT src="scripts/searchui.js"></SCRIPT>
<LINK href="style/SearchUI.css" type=text/css rel=stylesheet>
<LINK href="style/Flyout.css" type=text/css rel=stylesheet>
<LINK href="style/webparts.css" type=text/css rel=stylesheet>

<STYLE type=text/css>

	.clsPartHead
	{
   	 FONT-SIZE: 82%;
   	 CURSOR: move;
   	 COLOR: #ffffff;
   	 FONT-FAMILY: verdana, arial, helvetica;
   	 BACKGROUND-COLOR: #6699cc
	}
</STYLE>

</head>

<BODY  leftMargin=3 topMargin=0 marginwidth="0" marginheight="0">

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
	Response.Write ""
	Response.End 
End If


Set Conn = GetActiveConnection("KBroker")

initMenus()

'Outputs js script to hideall menus when a main menu is clicked
Function getScript(mnus())
    'Dim temp As String
    'Dim i As Integer
    temp = "<Script language=jscript>" & vbCrLf
    temp = temp & "function hideall(){" & vbCrLf
    For i = 0 To UBound(mnus)
        If mnus(i) = "" Then Exit For
        temp = temp & " document.all.item('" & "drp" & Mid(mnus(i), 4) & "').style.display=""none"";"
    Next
    temp = temp & "}" & vbCrLf & "</Script>"
    getScript = temp
End Function

Sub AddToArray(mnu)
    m_mnuArray(m_mnuCounter) = mnu
    m_mnuCounter = m_mnuCounter + 1
End Sub


sub drawMenu(id , caption , link , subMnu , mainMnu, child, imagePath)
	child = false
	code = ";domenu(this)"
	If subMnu Then
	    mnuClass = "NavOff"
	    moveCode = "mousemove(this)"
	    mouseOutCode = "mouseout(this)"
	    ClickCode = "mouseclick(this)"
	Else
	    code = ""
	    mnuClass = "SubMnuOff"
	    moveCode = "submove(this)"
	    mouseOutCode = "subout(this)"
	    ClickCode = "subclick(this)"
	End If
	If mainMnu Then code = ";doMainMenu(this)"
	If link = "#" Then
	    Action = "" 'caption
	Else
	    Action = "<a href='" & link & "' TARGET=maininfoR>"
	End If

	if mainMnu then
		%>
		<DIV class="tpl_flyout_container" style="WIDTH: 152px; MARGIN-LEFT: 0; MARGIN-TOP: 4px">
		<DIV class="ListNugget" id="foFlyoutMenuUp<%=id%>" style="WIDTH: 152px; MARGIN-LEFT: 0; MARGIN-TOP: 4px" name="foFlyoutMenuUp<%=id%>">
		<TABLE class="ListNuggetHeader" id="foFlyoutMenuUp<%=id%>Header" cellSpacing="0" cellPadding="0" width="150px" name="foFlyoutMenuUp<%=id%>Header">
		<TBODY>     
		<TR>     
		<TD  vAlign=left  class="clsPartHead"
		onselectstart="window.event.cancelBubble=true; return false;"      
		onclick="PartWrapperToggle('foFlyoutMenuUp<%=id%>');"><IMG class=clsPartHead src="images/gripblue.gif" height=19 width=15>
		<A class="ListNuggetTitle"	onclick="return PartWrapperToggle('foFlyoutMenuUp<%=id%>');"  
		href="javascript:PartWrapperToggle('foFlyoutMenuUp<%=id%>');"><%=caption%></A></TD>     
		<TD class="ListNuggetButtonCellBlue" onclick="PartWrapperToggle('foFlyoutMenuUp<%=id%>');" bgcolor="#6699cc">     
		<DIV class="ListNuggetButton">
			<IMG class="ListNuggetUpButton" id="foFlyoutMenuUp<%=id%>Up" height=17 alt="Hide options"      
			src="images/blue-chevron_up.gif" width="17" align="right"	border=0 name="foFlyoutMenuUp<%=id%>Up">						
			<IMG class="ListNuggetDownButton" id="foFlyoutMenuUp<%=id%>Down" height="17" alt="Options" src="images/gray-chevron_down.gif" width="17" align="right" border="0" name="foFlyoutMenuUp<%=id%>Down">			
		</DIV></TD></TR></TBODY></TABLE>
		<DIV class="ListNuggetBody" id="foFlyoutMenuUp<%=id%>Body" name="foFlyoutMenuUp<%=id%>Body" STYLE="width: 150px">     
		
		<!--LOAD THE PAGE WITH ALL THE MENUS CLOSED-->
		
		<script language="Javascript">
			//PartWrapperToggle('foFlyoutMenuUp<%=id%>', 'link')
		</script>
		
		<!-- FLYOUT MENU -->     
		<DIV WIDTH="150px" STYLE="MARGIN-LEFT: 0px">
		<%
	else
		
		if child then		
			showChild = " if (document.all.item('" & id & "').style.visibility == 'hidden') " & vbCrLf
			showChild = showChild & " document.all.item('" & id & "').style.visibility = ''" &  vbCrLf
			showChild = showChild & " else document.all.item('" & id & "').style.visibility = 'hidden' "			
			caption = caption & "<img src='images\closed.gif' align=right border=0>"		
		else
			hideChild = ""
			showChild = ""
		end if
		%>
		<TABLE name="test<%=id%>" STYLE="MARGIN-LEFT: 0px" id="test<%=id%>" cellSpacing=0 cellPadding=0 width="150px" border="0" onmouseover="javascript: Change(this); <%=showChild%> " onmouseout="javascript: UnChange(this); <%=showChild%> " onclick="">
        <TBODY>
			<TR>     
				<A href="javascript: window.parent.parent.frames['tabs'].document.all.item('ReportTitle').innerHTML = '<%=caption%>'; window.parent.frames['maininfoR'].location='<%=link%>'">
				<TD class="flyoutLink" handle="298" bgColor="#f0f8ff">
									 <%
					Response.Write caption

					if child then
						DoChildMenu id
					end if				
					%>
					
				</TD>
				</A>
			</TR>
		</TBODY>
		</TABLE>		
		<%
	end if	
	
	  
End sub

'This procedure checks for submenus to the main menus and outputs them to the browser
'It is similar to contents_respond above only that this function is called recursively for each instance of a submenu
Sub DoSubMenu(mnuID)  
    
    
    rsSubChild.Filter = "mainmenuID ='" & mnuID & "'"
   
    With Response       
        %>
		<TABLE class=flyoutMenu id=idFlyoutMenu cellSpacing=0 cellPadding=0 border=0>     		        
		        <TR>     
		          <TD>        
        <%
        While Not (rsSubChild.EOF Or rsSubChild.BOF)        
            
                Drp = rsSubChild.Fields("submenu").Value 
                If Not (rsSubChild.Fields("mnuaction").Value  = "") Then
                    Action = "/" & rsSubChild.Fields("mnuaction").Value 
                Else
                    Action = "#"
                End If    
                       
                If IsNull(rs("Image")) Then
					mnuImage = "images/linknote.gif"
				ElseIf rs("Image") = "" Then
					mnuImage = "images/linknote.gif"
				Else
					mnuImage = rs("Image")
				End If
                
                If Drp Then
					drawMenu rsSubChild.Fields("menuID").Value , rsSubChild.Fields("mnucaption").Value , Action, Drp, False, True, mnuImage
				else
					drawMenu rsSubChild.Fields("menuID").Value, rsSubChild.Fields("mnucaption").Value, Action, Drp, False, False, mnuImage
                End If
           
            rsSubChild.MoveNext
        Wend
       
		%>
		</TD></TR></TABLE></DIV><!-- END FLYOUT MENU --></DIV></DIV></DIV>
		<%
       
       
    End With    
   
End Sub


Sub DoChildMenu(mnuID)    
    
    exit sub
    
    rsChild.Filter = "mainmenuID = '" & mnuID & "'"
    
    With Response
        
        %>
		<DIV id="<%=mnuID%>" style="position: absolute; visibility: hidden; width: 150px; z-index: 1">
        <%If Not (rsChild.EOF Or rsChild.BOF) Then%>
        		<TABLE bgcolor="#C0C0C0" name="test<%= rsChild.Fields("menuID").Value %>" id="test<%= rsChild.Fields("menuID").Value  %>" cellSpacing=0 cellPadding=0 width="150px" border="0">
		        <TBODY>     
		<%
			While Not (rsChild.EOF Or rsChild.BOF)        
          
                Drp = rsChild.Fields("submenu").Value 
                If Not (rsChild.Fields("mnuaction").Value  = "") Then
                    Action = rsChild.Fields("mnuaction").Value 
                Else
                    Action = "#"
                End If                
                
                If IsNull(rs("Image")) Then
					mnuImage = "images/linksubnote.gif"
				ElseIf rs("Image") = "" Then
					mnuImage = "images/linksubnote.gif"
				Else
					mnuImage = rs("Image")
				End If
                   
				%>
					<TR  onmouseover="javascript: this.bgColor='GRAY'; window.status='';" onmouseout="javascript: this.bgColor='#C0C0C0'; window.status='';">     
						<A href="<%=Action%>">
							<TD class="flyoutLink" handle="298">
								<img src="<%= mnuImage %>" style="filter:alpha(opacity=50); -moz-opacity:0.5" width=20 height=20>&nbsp;<%= rsChild.Fields("mnucaption").value %>
							</TD>
						</A>
					</TR>
				<%             
				rsChild.MoveNext
			Wend%>
			
					
				</TBODY>
				</TABLE>		
		<%End If%>
		</DIV><!-- END FLYOUT MENU -->
		<%
    
       
    End With
  
End Sub

Sub initMenus()	
   
	dim SQL    
  
    mnuCatID = Request.QueryString("mnuID")
    userID = Session("UserID")
'IsMainMenu = 1 AND
	SQL = "SELECT * FROM Menus WHERE MainMenuID = " & mnuCatID & " AND IsReport = 1 AND  EXISTS(SELECT     MenuGroups.ID " & _
				" FROM         UserGroups INNER JOIN " & _
				"                      MenuGroups ON UserGroups.GroupID = MenuGroups.groupID " & _
				"			WHERE     (UserGroups.UserID = " & userID & ") AND (MenuGroups.MenuID = Menus.menuID))   ORDER BY mnuCaption"
		

	Set Rs = Conn.Execute(SQL)

	If Not (Rs.EOF Or Rs.BOF) Then
	Else
		Set Rs = Nothing
		Set Conn = Nothing
		Response.Write "There are currently no reports for the specified function."	
		Response.End
	End If	
			
    Set rsChild = Server.CreateObject("ADODB.Recordset")    
    Set rsSubChild = Server.CreateObject("ADODB.Recordset")    
    
    rsSubChild.CursorLocation = 3 
   	SQL = "SELECT * FROM Menus WHERE IsReport =  1 AND EXISTS(SELECT     MenuGroups.ID " & _
						" FROM         UserGroups INNER JOIN " & _
						"                      MenuGroups ON UserGroups.GroupID = MenuGroups.groupID " & _
						"			WHERE     (UserGroups.UserID = " & userID & ") AND (MenuGroups.MenuID = Menus.menuID))   ORDER BY mnuCaption"
	
	rsSubChild.Open SQL, Conn.ConnectionString, 1, 1 
    
    rsChild.CursorLocation = 3
    rsChild.Open SQL, Conn.ConnectionString, 1, 1 
   	
    With Response		        
        %>
		<TABLE name="CenterPanel" id="CenterPanel" cellSpacing=0 cellPadding=0 width="150px" marginLeft=0px border=0>
		  <TBODY>     
		  <TR>     
		    <TD vAlign=top align=left>        
        <%
        'drawMenu 0, "Reports" , "" , True , True, False, ""
        
        
        While Not (rs.EOF Or rs.BOF)
			    
            
					
					
                    Drp = rs("submenu") 'True or false
                    If Not (rs("mnuaction") = "") Then 'Action is a html link
                        Action = rs("mnuaction")
                    Else
                        Action = "#"
                    End If
                    
                    if Mid (rs("mnuCaption"), 1, 2) = "zz" then
						mnuCaption = Mid (rs("mnuCaption"), 3, len (rs("mnuCaption")) - 2)
					else
						mnuCaption = rs("mnuCaption")
                    end if                    
                    
                    If IsNull(rs("Image")) Then
						mnuImage = "images/linknote.gif"
					ElseIf rs("Image") = "" Then
						mnuImage = "images/linknote.gif"
					Else
						mnuImage = rs("Image")
					End If
                    
                    
                    If Rs.Fields("IsMainMenu") = 1 Then
                    
						drawMenu  rs("menuID"), rs("mnuCaption") , "" , True , True, False, ""
                    Else
						drawMenu rs("menuID"), mnuCaption, Action, Drp, False, True, mnuImage
                    End If
                    
                    AddToArray rs("menuID") 'Remember the id of each menu sent to the browser
                    
                    If Drp Then 'True if the menu has a submenu
                        DoSubMenu rs("menuID")                        
                    End If					                    
            
            rs.MoveNext
        Wend
      
       %>
		           </TD>     
		    </CENTER>     
		  </TR>
		  </TBODY>
		</TABLE> 
       <%        
        .Write getScript(m_mnuArray)        
    End With
    
    'destroy objects
    Set rs = Nothing    
    Set rsChild = Nothing
    Set rsSubChild = Nothing
    
    
End Sub
%>

<SCRIPT LANGUAGE="JavaScript">
	function Change(row){
		row.bgColor = "white"		
		row.style.borderstyle = "outset"
		row.style.bordercolor = "#B0B0B0"
		row.style.borderwidth = "1"		
		//document.all.item("IMG" + row.name).filters.alpha.opacity=100;		

	}
	
	function UnChange(row){
		row.bgColor = "#f0f8ff"		
		row.style.borderstyle = ""
		row.style.bordercolor = ""
		row.style.borderwidth = "0"
		//document.all.item("IMG" + row.name).filters.alpha.opacity=50;
	}

</Script>
