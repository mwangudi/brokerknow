<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EditAccountType"
	const DataEntity = "AccountType"
	const DataEntityPlural = "Account Types"
	const ActionFolder = "Data"
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim rsEdit
	Dim guid
	Dim guidStr
	
	action = ucase(Request.Form("action"))
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
				Dim parentType
				Dim Setting
       
				parentType = Request.Form("txtParentType")
				if parentType = "n/a" then
						parentType = "NULL"
				else
						parentType = Request.Form("cboParentType") 
				end if
				Setting = Request.Form("txtSetting")		        
				if parentType <> "NULL" then
						'validate Parent Type
						If Trim(parentType) = "" Then%>
						         <script language = 'vbscript'>
						         		ShowMessage "Please specify the Parent Type"
						        </script>
						         <% response.end
						End If
				end if
				
				'validate Setting
				If Trim(Setting) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Setting"
				        </script>
				         <% response.end
				End If
				'validate size of Setting
				If Len(Setting) > 200 Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Setting can only be 200 characters in length"
				        </script>
				         <% response.end
				End If 
				
				'save data		
				Set conn = GetActiveConnection("KBroker")
						
				sqlStr = "UPDATE [AccountType] SET " & _
						"		AccountTypeName = " & "'" & Replace(Setting, "'", "''") & "'" & "" & _
						"       ,AccountTypeParent = " & " " & parentType & " " & _
						"        WHERE AccountType_DPA_  = " & ID
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

<script >
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
		
		
</script>
</head>
<%
Set conn = GetActiveConnection("KBroker")
sqlStr = "SELECT * FROM AccountTypeList WHERE AccountType_DPA_=" & ID
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		window.self.ShowMessage "The selected <%=DataEntity%> cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
%>
<body Class="Dialog">

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
 
	<%if rs.Fields("AccountTypeLevel_DPA_") = 1 then%>
			<input type = 'hidden' name ='txtparentType' id = 'txtparentType' value='n/a'>
	<%else%>
		<tr>
		    <td >Parent Type</td>
		    <td ><select name = 'cboParentType' id = 'cboParentType' size="1">
		<%		
				Dim levelFilter
				
		        sqlStr = "SELECT * FROM [AccountTypeList] WHERE AccountTypeLevel_DPA_ IN (1,2) Order By AccountTypeLevel_DPA_,AccountTypeSetting"
		        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		        If Not (rsEdit.EOF Or rsEdit.BOF) Then
		                rsEdit.MoveFirst
		                levelFilter = 1
		                Do Until rsEdit.EOF
		                		if levelFilter <> rsEdit.Fields("AccountTypeLevel_DPA_") then
										levelFilter = rsEdit.Fields("AccountTypeLevel_DPA_")%>
										<option value = ''>**********</option>
								<%else
										ShownAccount=rsEdit.Fields("AccountTypeSetting") & ":" & rsEdit.Fields("AccountTypeParent")
										if rsEdit.Fields("AccountTypeSetting") = rs.Fields("AccountTypeParent") Then%>
		                						<option selected value = '<%=rsEdit.Fields("AccountType_DPA_")%>'><%=ShownAccount%></option>
		                				<%else%>
												<option value = '<%=rsEdit.Fields("AccountType_DPA_")%>'><%=ShownAccount%></option>
										<%end if
										rsEdit.MoveNext
								end if
								
		                Loop
		        End If
		%>

		    </select></td>
		    <td >

			</td>
		 </tr>
 <%end if%> 
  
  
   <tr>
    <td >Setting</td>
    <td ><input type = 'text' name ='txtSetting' id = 'txtSetting' size="20" value = '<%=rs.Fields("AccountTypeSetting")%>'></td>
    <td >

	</td>
  </tr>
  <tr>
	  <td width="100%" colspan=3 align="right" valign=absBottom>
		<BR><BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "AllowedNavigation()">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%= Rs.Fields("AccountType_DPA_").Value %>">
	</td>
  </tr>
</table>

</form>
</body>

</html>
