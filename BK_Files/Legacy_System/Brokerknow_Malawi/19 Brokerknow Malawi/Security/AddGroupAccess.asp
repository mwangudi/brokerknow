<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Group Access</title>
  <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 

 <script language='javascript'>
		function forceSubmit()
		{
			setOpener();
			document.frmAddGroupAccess.method='post';
			document.frmAddGroupAccess.target='_self';
			document.frmAddGroupAccess.submit();		
		}		
		function setOpener()
		{
			//window.self.opener = window.dialogArguments.opener;					
		}
</script>
</head>

<!--#include file="../libroutines.asp"-->

<%
	
   Dim action
   Dim conn 
   Dim sqlStr
   Dim rs
   Const dataMark = "~"
   
	
	action = ucase(Request.Form("action"))
	
	Set conn = GetActiveConnection("KBroker")
	  
	Set menuRs = Conn.Execute("SELECT * FROM Menus WHERE IsMainMenu = 1 ORDER BY mnuCaption")
	SQL = "SELECT * FROM Groups ORDER BY GroupName"
	Set GroupRs = Conn.Execute(SQL)
	
	If action = "EXECUTE" Then	

		
		If Not (GroupRs.EOF OR GroupRs.BOF) Then
			Do Until GroupRs.EOF			
				thisGroupID = GroupRs("GroupID")
				hasRights = False
				
				If Not (menuRs.EOF Or menuRs.BOF) Then
					Do Until menuRs.EOF   
		      		
			      		MenuID = menuRs("MenuID")
						withRights = Request.Form("RightsMove" & MenuID)
						
						If withRights <> "" Then
							If InStr(1, withRights, dataMark & thisGroupID & dataMark) > 0 Then
								hasRights = True
							Else
								hasRights = False
							End If										
						Else
							hasRights = False
						End If  
						
						If hasRights = True Then
							Set chkRs = Conn.Execute ("SELECT * FROM MenuGroups WHERE GroupID = " & thisGroupID & " AND MenuID = " & MenuID)
							If Not (chkRs.EOF Or chkRs.BOF) Then
								'ignore
							Else
								'update								
								Conn.Execute ("INSERT INTO MenuGroups (GroupID, MenuID) VALUES (" & thisGroupID & ", " & MenuID & ")")
								'children
								childSQL = "INSERT INTO MenuGroups (GroupID, MenuID) SELECT " & thisGroupID & " AS GroupID, MenuID FROM Menus WHERE (MainMenuID = " & MenuID & ") Or (MainMenuID IN (SELECT MenuID FROM Menus WHERE MainMenuID = " & MenuID & "))"
								Conn.Execute childSQL
							End If	
						Else
							delChildrenSQL = "DELETE FROM MenuGroups WHERE (GroupID = " & thisGroupID & ") AND (MenuID IN (SELECT MenuID FROM Menus WHERE MainMenuID = " & MenuID & ")  Or MenuID IN (SELECT MenuID FROM Menus WHERE MainMenuID = " & MenuID & "))"
							Conn.Execute (delChildrenSQL)
							Conn.Execute ("DELETE FROM MenuGroups WHERE GroupID = " & thisGroupID & " AND MenuID = " & MenuID)
						End If							
						    		
						menuRs.MoveNext
					Loop							
					menuRs.MoveFirst
				End If	
				
			GroupRs.MoveNext
			Loop
		End If		
		
        Set conn = Nothing
        WriteFraEnabledDialogCloseScript
        Response.End
   	End If  	
   	
   	
%>

<body Class="Dialog" OnLoad="setOpener();"  OnLoad=" VBScript: DoResizeWin">
<form name = 'frmAddGroupAccess' method = 'post' action = 'AddGroupAccess.asp'>

<table border="0" cellspacing="1" cellpadding="2" id="mainTable">
	<tr>
	<td><b><font size="2" face="Verdana" color="#0000FF">Groups:</font></b></td>
	<td colspan="2" rowspan="2" valign="Top">
	<%
	
	If Not (menuRs.EOF Or menuRs.BOF) Then
		
		Do Until menuRs.EOF
		
			Caption = Trim(menuRs("mnuCaption"))
			menuID = menuRs("MenuID")
			%>      
			<table border="0" width="100%" cellspacing="1" CellPadding="1" Class="tdRule">
			<tr>
			<td width="200%" colspan="2"><font size="1" face="Verdana" color="#000080"><b><%= Caption %>:</b></font></td>
			</tr>
			<tr>
			<td width="100%"><input type="button" value=" &lt; " name="Move<%= menuID %>" Class=Buttons OnClick="JavaScript: Move(this)">&nbsp;&nbsp;&nbsp;<input type="button" value=" &gt; " name="Move<%= menuID %>" Class=Buttons OnClick="JavaScript: Move(this)"></td>
			<td width="100%">&nbsp;&nbsp;
			<select size="2" name="RightsMove<%= menuID %>" multiple  OnKeyPress="JavaScript: if (event.keyCode==46) Move(this)">
			<%
			
			SQL = "SELECT Groups.* FROM Groups INNER JOIN MenuGroups ON Groups.GroupID = MenuGroups.GroupID WHERE MenuGroups.MenuID = " & menuID
			Set rightsRs = Conn.Execute(SQL)
			If Not (rightsRs.BOF Or rightsRs.EOF) Then
				Do Until rightsRs.EOF
					captionName = rightsRs("groupName")
					%>
					<Option Value="<%= dataMark & rightsRs("GroupID") & dataMark %>"><%= captionName %></Option>
					<%
					rightsRs.MoveNext
				Loop
			End If
			rightsRs.Close
			Set rightsRs = Nothing%>
			</select></td>
			</tr>
			</table>               
			<%
			menuRs.MoveNext
		Loop
	End If
			
	menuRs.Close
	Set menuRs = Nothing%>
	</td>
	</tr>
	<tr valign="Top">
	
	<td nowrap ID="GroupsTD">
	<%
	If Not (GroupRs.BOF Or GroupRs.EOF) Then%>
		<select size="1" name="Groups" multiple>
		<option>Select a group to add</option>
		<%
		Do Until GroupRs.EOF
			gName = GroupRs("GroupName")
			%>
			<Option Value="<%= dataMark & GroupRs("GroupID") & dataMark %>"><%= gName %></Option>
			<%
			GroupRs.MoveNext
		Loop
		%>
		</select>
		<%
		Else
		Response.Write "No Current Groups"
		Response.End
	End If
	Set GroupRs = Nothing
	Set Conn = Nothing %></td>
	        
  </tr>
  
  <tr>
   <td width="100%" colspan="3" align=right>
		<BR>	
		<input type = 'submit' Class='Buttons' name ='cmdAdd' id = 'cmdAdd' value=" Save "  OnClick="VBScript: SelectForm">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
		<input type = 'hidden' name ='buttonAction' id = 'buttonAction' value="Save">
      </td>
  </tr>

</table>
</form>
<!--JAVA CODE-->
<Script Language="JavaScript">
<!--Begin
	document.all.item("Groups").size = Math.round(document.all.item("GroupsTD").clientHeight / 10) ;
//===========BEGIN MOVE FUNCTION ============================= 
function Move(Btn){

var todo = Btn.value;
var Users = document.all.item("Groups");
var loop;

if (todo.search(">")>0){
 if (Users.selectedIndex==-1) return(ShowMessage("Select a group from the groups list."))
  var InsertList = document.all.item("Rights" + Btn.name) 
  AddOption(Users,InsertList) 
  }
else{
 RemoveOption(document.all.item("Rights" + Btn.name))
 }
}

//=========END FUNCTION ====================================================

//=========BEGIN DROP-DOWN SELECT FUNCTION FOR FORM POSTING======

 function SelectAll(Object){
 //select all upwards
  for (loop=Object.length-1; loop>-1;
   loop--)
    {
     Object.options[loop].selected = true
    }
  }

//============END SELECT FUNCTION===================================

//==========BEGIN REMOVE OPTION/S FROM DROP-DOWN FUNCTION ON THE FLY=====
  function RemoveOption(Field){
	Selection = new Boolean();
	if (Field.length==0) return(ShowMessage("The list is empty."))
	for (loop=Field.length - 1; loop >= 0; loop--) {
	    var GoneOption = Field.options[loop]
		if (GoneOption.selected==true) {
	      		Selection = true;
	      		Field.remove(GoneOption.index);
	      }
	    }
	    
   if (Selection==false) ShowMessage("Select a group to remove from the List.")
   
  }
//==============END REMOVE OPTION/S FUNCTION====================

//=========BEGIN ADD OPTION TO DROP-DOWN ON THE FLY=============

  function AddOption(Input,Output){    
    for (loop=0; loop < Input.length; loop++){
    		if (Input.options[loop].selected && loop !== 0){
    		    NewOption = new Option();   			    
			    NewOption.text = Input.options[loop].text;
			    NewOption.value = Input.options[loop].value;			
			    NewOption.selected = false;
			    if (!CheckDuplicates(Output, NewOption.value)) Output.add(NewOption, 0)
    		}
    		
    }
    
  }

function CheckDuplicates(DupPut, valText){
	var loop;
   for (loop=0; loop < DupPut.length;loop++){
      if (DupPut.options[loop].value==valText){
	       	return(true) ;
       }
     }
 }   
   
  
//========END ON THE FLY ADD OPTION FUNCTION=====================

-->
</Script>
 <!--END CODE-->
<Script Language="VBScript">
	Function SelectForm
		For Each Thing In frmAddGroupAccess
			If InStr(1, Thing.Name, "RightsMove") > 0 Then
				SelectAll Thing
			End If
		Next
	End Function
	
	Function DoResizeWin	
		Dim nHeight, nWidth
		Dim defMaxHeight, defMaxWidth
		
		defMaxHeight = (screen.availHeight) - 100
		defMaxWidth = (screen.availWidth) - 100
		
		nHeight = (document.all.item("mainTable").clientHeight) + 50
		nWidth = (document.all.item("mainTable").clientWidth) + 50
		
		If nHeight > defMaxHeight Then
			nHeight = defMaxHeight
		End If
		
		If nWidth > defMaxWidth Then	
			nWidth = defMaxWidth
		End If		
				
		window.parent.dialogHeight = nHeight & "px"
		window.parent.dialogWidth = nWidth & "px"
		
		
	End Function
</Script>
</body>

