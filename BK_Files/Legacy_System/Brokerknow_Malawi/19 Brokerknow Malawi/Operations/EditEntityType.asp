<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EditEntityType"
	const DataEntity = "EntityType"
	const DataEntityPlural = "EntityTypes"
	const ActionFolder = "Operations"
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim rsEdit
	Dim guid
	Dim guidStr
	
	action = ucase(Request.Form("action"))
	Level1 = Request.Form("cboAccountTypeLevel1")
	Level2 = Request.Form("cboAccountTypeLevel2")
	Level3 = Request.Form("cboAccountTypeLevel3")
	Code = Request.Form("txtEntTypeCode")
	EntTypeName = Request.Form("txtEntTypeName")
	
	ID = Request("ID")

	If Trim(ID) = "" Then%>
            <script language = 'vbscript'>
            		ShowMessage "No record specified for editing"
                		
            </script>
            <% response.end
    End If
        
	if action = "EXECUTE" then
		Dim buttonAction
		Dim reloadRequired
		
		reloadRequired = false
		buttonAction = Trim(Ucase(Request.Form("cmdAdd")))
		if buttonAction = "SAVE" then
				'Dim EntTypeName
				'Dim code
		        
				'EntTypeName = Request.Form("txtEntTypeName")
				'code = Request.Form("txtEntTypeCode")
		       
				 
				'validate Account Type
				If Trim(Level1) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Account Type"
				         		
				         </script>
				         <% response.end
				End If
				
				'validate Name
				If Trim(EntTypeName) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Name"
				         		
				         </script>
				         <% response.end
				End If
				'validate size of Name
				If Len(EntTypeName) > 100 Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Name can only be 100 characters in length"
				         		
				         </script>
				         <% response.end
				End If
				'validate Code
				If Trim(Code) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Code"
				         		
				         </script>
				         <% response.end
				End If
				'validate size of Code
				If Len(EntTypeName) > 50 Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Code can only be 50 characters in length"
				         		
				         </script>
				         <% response.end
				End If
				
				if Level2 = "" then
						Level2 = "Null"
				end if
							
				if Level3 = "" then
						Level3 = "Null"
				end if
				
				'save data		
				Set conn = GetActiveConnection("KBroker")
				
				sqlStr = "UPDATE [EntityType] SET " & _
						"	EntityTypeCode = " & "'" & code & "'" & _
						"	,EntityTypeName = " & "'" & EntTypeName & "'" & _
						"	,AccountType_DPA_ = " & " " & Level1 & " " & _
						"	,AccountType_DPA_2 = " & " " & Level2 & " " & _
						"	,AccountType_DPA_3 = " & " " & Level3 & " " & _
						"	WHERE EntityType_DPA_  = " & ID
						
				sqlStr = SQLServerFormat(HandleQuote(sqlStr))


				conn.Execute sqlStr
				
				conn.Close
				Set conn = Nothing
				WritefraEnabledDialogCloseScript
				Response.End
			end if
			Dim clientCode
        
			clientCode = "var validNavigate = true;" & chr(13)
			%>
			<script>
				<%=clientCode%>
			</script>
			<%
			response.End
	elseif action = "FETCH_ACCOUNT_TYPES" then
			'do nothing
			Set conn = GetActiveConnection("KBroker")
	else
			Set conn = GetActiveConnection("KBroker")
			sqlStr = "SELECT * FROM EntityType WHERE EntityType_DPA_=" & ID
			        
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If rs.EOF Or rs.BOF Then%>
			        <script language = 'vbscript'>
			        		window.self.ShowMessage "The selected <%=DataEntity%> cannot be retrieved for editing"
			        		window.self.close
			        </script>
			        <% response.end
			End If
			
			Level1 = rs.Fields("AccountType_DPA_").value
			if isnull(rs.Fields("AccountType_DPA_2").value) then
					Level2 = ""
			else
					Level2 = rs.Fields("AccountType_DPA_2").value
			end if
			
			if isnull(rs.Fields("AccountType_DPA_3").value) then
					Level3 = ""
			else
					Level3 = rs.Fields("AccountType_DPA_3").value
			end if
			
			EntTypeName = rs.Fields("EntityTypeName").value
			Code = rs.Fields("EntityTypeCode").value
   	end If
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

<script language='vbscript'>

			function LevelSelected()
 					frm<%=DataSource%>.elements("action").value = "Fetch_Account_Types"
 					frm<%=DataSource%>.submit
 							
			end function
</script>
<script language='javascript'>
		var validNavigate = false;
		function ReleaseRecord()
		{
			if(!validNavigate)
			{
 				event.returnValue = "Please use the cancel button to close the dialog"
 			}
		}
		
		function AllowedNavigation()
		{
			validNavigate = true;
		}
		
		function FetchLevel()
		{
			document.frmMain.target = "_self"
			LevelSelected();
			
		}

</script>
</head>
<%
'Set conn = GetActiveConnection("KBroker")
'sqlStr = "SELECT * FROM EntityType WHERE SystemMaintained = 0 AND EntityType_DPA_=" & ID
'sqlStr = "SELECT * FROM EntityType WHERE EntityType_DPA_=" & ID
        
        'Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        'If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		'window.self.ShowMessage "The selected <%=DataEntity%> cannot be retrieved for editing"
                		
                </script>
                <% 'response.end
        'End If
%>
<body Class="Dialog">

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
 
  <tr>
    <td >Type 1</td>
    <td ><select name = 'cboAccountTypeLevel1' id = 'cboAccountTypeLevel1' size="1" onchange='FetchLevel()'>
<%
        
        sqlStr = "SELECT * FROM [AccountTypeLevel1] Order By AccountTypeName"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
						if rsEdit.Fields("AccountType_DPA_").value = cint(Level1) then%>
										<option selected value = '<%=rsEdit.Fields("AccountType_DPA_")%>'><%=rsEdit.Fields("AccountTypeName")%></option>
								<%else%>
										<option value = '<%=rsEdit.Fields("AccountType_DPA_")%>'><%=rsEdit.Fields("AccountTypeName")%></option>
								<%end if
                        rsEdit.MoveNext
                Loop
        End If
%>

    </select></td>
    <td >

	</td>
  </tr>
  <tr>
    <td >Type 2</td>
    <td ><select name = 'cboAccountTypeLevel2' id = 'cboAccountTypeLevel2' size="1" onchange='FetchLevel()'>
    <option selected value =''></option>
<%
        sqlStr = "SELECT * FROM [AccountTypeLevel2] WHERE AccountTypeParent = " &  Level1 & " Order By AccountTypeName"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                        if Level2 = "" then
								if cbool(rsEdit.Fields("DefaultSelection")) then%>
										<option selected value = '<%=rsEdit.Fields("AccountType_DPA_")%>'><%=rsEdit.Fields("AccountTypeName")%></option>
								<%else%>
										<option value = '<%=rsEdit.Fields("AccountType_DPA_")%>'><%=rsEdit.Fields("AccountTypeName")%></option>
								<%end if
						else
								if rsEdit.Fields("AccountType_DPA_").value = cint(Level2) then%>
										<option selected value = '<%=rsEdit.Fields("AccountType_DPA_")%>'><%=rsEdit.Fields("AccountTypeName")%></option>
								<%else%>
										<option value = '<%=rsEdit.Fields("AccountType_DPA_")%>'><%=rsEdit.Fields("AccountTypeName")%></option>
								<%end if
						end if
                        rsEdit.MoveNext
                Loop
        else
				Level2 = ""
        End If
       
%>
	
    </select></td>
    <td >
    
	</td>
  </tr>
  <tr>
    <td >Type 3</td>
    <td ><select name = 'cboAccountTypeLevel3' id = 'cboAccountTypeLevel3' size="1" onchange='FetchLevel()'>
    <option selected value =''></option>
<%
		If Level2 = "" then
				Level2 = 0
		end if
		
        sqlStr = "SELECT * FROM [AccountTypeLevel3] WHERE AccountTypeParent = " &  Level2 & " Order By AccountTypeName"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                        if Level3 = "" or Level2 = 0 then
								if cbool(rsEdit.Fields("DefaultSelection")) then%>
										<option selected value = '<%=rsEdit.Fields("AccountType_DPA_")%>'><%=rsEdit.Fields("AccountTypeName")%></option>
								<%else%>
										<option value = '<%=rsEdit.Fields("AccountType_DPA_")%>'><%=rsEdit.Fields("AccountTypeName")%></option>
								<%end if
						else
								if rsEdit.Fields("AccountType_DPA_").value = cint(Level3) then%>
										<option selected value = '<%=rsEdit.Fields("AccountType_DPA_")%>'><%=rsEdit.Fields("AccountTypeName")%></option>
								<%else%>
										<option value = '<%=rsEdit.Fields("AccountType_DPA_")%>'><%=rsEdit.Fields("AccountTypeName")%></option>
								<%end if
						end if
                        rsEdit.MoveNext
                Loop
        else
				If Level3 <> "00" then
						Level3 = ""
				end if
        End If
       
%>
	
    </select></td>
    <td >
    
	</td>
  </tr>
  <tr>
    <td >Name</td>
    <td ><input type = 'text' name ='txtEntTypeName' id = 'txtEntTypeName' size="20" value = '<%=EntTypeName%>'></td>
    <td >

	</td>
  </tr>
  <tr>
    <td >Code</td>
    <td ><input type = 'text' name ='txtEntTypeCode' id = 'txtEntTypeCode' size="20" value = '<%=Code%>'></td>
    <td >

	</td>
  </tr>
  <tr>
	  <td width="100%" colspan=3 align="right" valign=absBottom>
		<BR><BR>
		<%'if Rs.Fields("SystemMaintained").Value = 0 then%>
				<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "AllowedNavigation()">
		<%'end if%>
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%= ID %>">
	</td>
  </tr>
</table>

</form>
</body>

</html>
