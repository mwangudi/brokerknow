<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EditGenericSetting"
	const DataEntity = "GenericSetting"
	const DataEntityPlural = "Generic Settings"
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
				Dim EntType
				Dim Generic
				Dim Setting
       
				EntType = Request.Form("cboEntityType") 
				Generic = Request.Form("cboGeneric")
				Setting = Request.Form("txtSetting")
		       
				'validate Entity Type
				If Trim(EntType) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Entity Type"
				        </script>
				         <% response.end
				End If
				'validate Generic
				If Trim(Generic) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Generic"
				        </script>
				         <% response.end
				End If
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
						
				sqlStr = "UPDATE [GenericSetting] SET " & _
						"		GenericSettingDescription = " & "'" & Replace(Setting, "'", "''") & "'" & "" & _
						"       ,EntityType_DPA_ = " & " " & EntType & " " & _
						"		,Generic_DPA_ = " & " " & Generic & " " & "" & _
						"        WHERE GenericSetting_DPA_  = " & ID
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
sqlStr = "SELECT * FROM GenericSetting WHERE GenericSetting_DPA_=" & ID
        
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
 
  <tr>
    <td >Entity Type</td>
    <td ><select name = 'cboEntityType' id = 'cboEntityType' size="1">
<%		
        sqlStr = "SELECT * FROM [FullEntityTypeList] Order By EntityTypeName"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("EntityType_DPA_") = rs.Fields("EntityType_DPA_") Then%>
                			<option selected value = '<%=rsEdit.Fields("EntityType_DPA_")%>'><%=rsEdit.Fields("EntityTypeName")%></option>
                		<%else%>
                        <option value = '<%=rsEdit.Fields("EntityType_DPA_")%>'><%=rsEdit.Fields("EntityTypeName")%></option>
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
    <td >Generic</td>
    <td ><select name = 'cboGeneric' id = 'cboGeneric' size="1">
<%
        sqlStr = "SELECT * FROM [GenericList] Order By GenericDescription"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("Generic_DPA_") = rs.Fields("Generic_DPA_") Then%>
                			<option selected value = '<%=rsEdit.Fields("Generic_DPA_")%>'><%=rsEdit.Fields("GenericDescription")%></option>
                		<%else%>
                        <option value = '<%=rsEdit.Fields("Generic_DPA_")%>'><%=rsEdit.Fields("GenericDescription")%></option>
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
    <td >Setting</td>
    <td ><input type = 'text' name ='txtSetting' id = 'txtSetting' size="20" value = '<%=rs.Fields("GenericSettingDescription")%>'></td>
    <td >

	</td>
  </tr>
  
  <tr>
	  <td width="100%" colspan=3 align="right" valign=absBottom>
		<BR><BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "AllowedNavigation()">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%= Rs.Fields("GenericSetting_DPA_").Value %>">
	</td>
  </tr>
</table>

</form>
</body>

</html>
