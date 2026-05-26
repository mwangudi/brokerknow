<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Group Member/s</title>
  <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
</head>

<!--#include file="../libroutines.asp"-->

<%
	
   Dim action
   Dim conn 
   Dim sqlStr
   Dim rs
   Const dataMark = "~"
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")
	
	If ID = "" Then%>
		<Script Language="JavaScript">
			ShowMessage("No record specified for editing")
		</Script>
		<%WriteDialogRefuseOpenScript
		Response.End
	End If
	
	Set conn = GetActiveConnection("KBroker")
	Set GroupRs = Conn.Execute("SELECT * FROM Groups WHERE GroupID = (SELECT GroupID FROM UserGroups WHERE MemberID = " & ID & ")")    
	
	If action = "EXECUTE" then
		Set StaffRs = Conn.Execute("SELECT * FROM Users")
		If Not (StaffRs.EOF OR StaffRs.BOF) Then
			Do Until StaffRs.EOF
				thisStaffID = StaffRs("UserID")
				hasRights = False
				If Not (groupRs.EOF Or groupRs.BOF) Then
						groupID = groupRs("GroupID")
						withRights = Request.Form("RightsMove" & groupID)
						If withRights <> "" Then
							If InStr(1, withRights, dataMark & thisStaffID & dataMark) > 0 Then
								hasRights = True
							Else
								hasRights = False
							End If										
						Else
							hasRights = False
						End If  
						
						If hasRights = True Then
							Set chkRs = Conn.Execute ("SELECT * FROM UserGroups WHERE GroupID = " & groupID & " AND UserID = " & thisStaffID)
							If Not (chkRs.EOF Or chkRs.BOF) Then
								'ignore
							Else
								'update
								Conn.Execute ("INSERT INTO UserGroups (GroupID, UserID) VALUES (" & groupID & ", " & thisStaffID & ")")
							End If	
						Else
							Conn.Execute ("DELETE FROM UserGroups WHERE GroupID = " & groupID & " AND UserID = " & thisStaffID)
						End If					
						   
				End If	
				
			StaffRs.MoveNext
			Loop
		End If		
		
        Set conn = Nothing
        WriteFraEnabledDialogCloseScript
        Response.End
   	End If  	
   	
   	
%>



<body Class="Dialog" OnLoad="VBScript: DoResizeWin">

<form name = 'frmEditGroupMembers' method = 'post' action = 'EditGroupMembers.asp'>

<table border="0" cellspacing="1" cellpadding="2" id="mainTable">
	<tr>
	<td><b><font size="2" face="Verdana" color="#0000FF">Users:</font></b></td>
	<td colspan="2" rowspan="2">
	<%
	'begin  
	
	If Not (GroupRs.EOF Or GroupRs.BOF) Then
		
		
			Caption = GroupRs("GroupName")
			groupID = GroupRs("GroupID")
			%>      
			<table border="0" width="100%" cellspacing="1" CellPadding="1">
			<tr>
			<td  colspan="2"><font size="1" face="Verdana" color="#000080"><b><%= Caption %>:</b></font></td>
			</tr>
			<tr>
			<td><input type="button" value=" &lt; " name="Move<%= groupID %>" Class=Buttons OnClick="JavaScript: Move(this)">&nbsp;&nbsp;&nbsp;<input type="button" value=" &gt; " name="Move<%= groupID %>" Class=Buttons OnClick="JavaScript: Move(this)"></td>
			<td>&nbsp;&nbsp;
			<select size="10" name="RightsMove<%= groupID %>" id="groupSelect" multiple Class=Selects OnKeyPress="JavaScript: if (event.keyCode==46) Move(this)">
			<%
			
			SQL = "SELECT Users.* FROM Users INNER JOIN UserGroups ON Users.UserID = UserGroups.UserID WHERE UserGroups.GroupID = " & groupID
			Set rightsRs = Conn.Execute(SQL)
			If Not (rightsRs.BOF Or rightsRs.EOF) Then
				Do Until rightsRs.EOF
					uName = rightsRs("OtherNames") & " " & rightsRs("Surname")
					%>
					<Option Value="<%= dataMark & rightsRs("UserID") & dataMark %>"><%= uName %></Option>
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
			
	End If
			
	GroupRs.Close
	Set GroupRs = Nothing%>
	</td>
	</tr>
	<tr valign="Top">
	<%SQL = "SELECT * FROM Users ORDER BY OtherNames, SurName"
	Set Users = Conn.Execute(SQL)%>
	<td nowrap ID="userTD">
	<%
	If Not (Users.BOF Or Users.EOF) Then%>
		<select size="10" name="Users" multiple  Class=Selects>
		<option>Select a user to add</option>
		<%
		Do Until Users.EOF
			User = Users("OtherNames") & " " & Users("SurName")
			%>
			<Option Value="<%= dataMark & Users("UserID") & dataMark %>"><%= User %></Option>
			<%
			Users.MoveNext
		Loop
		%>
		</select>
		<%
		Else
		Response.Write "No Current Users"
		Response.End
	End If
	Set Users = Nothing
	Set Conn = Nothing %></td>
	        
  </tr>
  
  <tr>
   <td width="100%" colspan="3" align=right>
		<BR>
	
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " OnClick="VBScript: SelectForm">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">		
		<input type = 'hidden' name ='ID' id = 'ID' value="<%= ID %>">&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
      </td>
  </tr>

</table>
</form>
<!--JAVA CODE-->
<Script Language="JavaScript">
<!--Begin
	
//===========BEGIN MOVE FUNCTION FOR SELECTED STAFF============================= 
function Move(Btn){

var todo = Btn.value;
var Users = document.all.item("Users");
var loop;

if (todo.search(">")>0){
 if (Users.selectedIndex==-1) return(ShowMessage("Select a user from the users list."))
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
	    
   if (Selection==false) ShowMessage("Select a user to remove from the List.")
   
  }
//==============END REMOVE OPTION/S FUNCTION====================

//=========BEGIN ADD OPTION TO DROP-DOWN ON THE FLY=============

  function AddOption(Input,Output){    
    for (loop=0; loop < Input.length; loop++){
    		if (Input.options[loop].selected && loop !== 0){
    		    NewOption = new Option();   			    
			    NewOption.text = Input.options[loop].text;
			    NewOption.value = Input.options[loop].value;			
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
		For Each Thing In frmEditGroupMembers
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

