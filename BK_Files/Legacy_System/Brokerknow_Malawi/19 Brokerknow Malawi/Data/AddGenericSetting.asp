<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "AddGenericSetting"
	const DataEntity = "GenericSetting"
	const DataEntityPlural = "Generic Settings"
	const ActionFolder = "Data"
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim guidStr 
	Dim guid 
		
	action = ucase(Request.Form("action"))
	
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
				set guid = Server.CreateObject("NDUtils.CGUID")
				guidStr = guid.GenerateGUID
				
				sqlStr = "INSERT INTO [GenericSetting] (GenericSettingDescription,GenericSetting_DPA_,GenericSetting_EIT_" & _
						"       ,EntityType_DPA_,Generic_DPA_) SELECT " & "'" & Replace(Setting, "'", "''") & "'" & " as GenericSettingDescription" & _
						"       ," & " " & "iif(isnull(max([GenericSetting_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'GenericSetting'),max([GenericSetting_DPA_]) + 1)" & " " & " as GenericSetting_DPA_" & _
						"       ," & "'" & guidStr & "'" & " as GenericSetting_EIT_," & " " & EntType & " " & " as EntityType_DPA_" & _
						"       ," & " " & Generic & " " & " as Generic_DPA_" & _
						"        FROM [GenericSetting]"
				Set conn = GetActiveConnection("KBroker")
				
				sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

				conn.BeginTrans
						conn.Execute sqlStr
				conn.CommitTrans
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
<title>Add <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

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
		
		

</script>
</head>

<body Class="Dialog">

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
  <tr>
    <td >Entity Type</td>
    <td ><select name = 'cboEntityType' id = 'cboEntityType' size="1">
<%
		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [FullEntityTypeList] Order By EntityTypeName"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                        if cbool(rs.Fields("DefaultSelection")) then%>
								<option selected value = '<%=rs.Fields("EntityType_DPA_")%>'><%=rs.Fields("EntityTypeName")%></option>
                        <%else%>
								<option value = '<%=rs.Fields("EntityType_DPA_")%>'><%=rs.Fields("EntityTypeName")%></option>
                        <%end if
                        rs.MoveNext
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
					<option value = ''></option>
<%
        sqlStr = "SELECT * FROM [GenericList] Order By GenericDescription"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF%>
						<option value = '<%=rs.Fields("Generic_DPA_")%>'><%=rs.Fields("GenericDescription")%></option>
                        <%rs.MoveNext
                Loop
        End If
%>

    </select></td>
    <td >

	</td>
  </tr>
  <tr>
    <td >Setting</td>
    <td ><input type = 'text' name ='txtSetting' id = 'txtSetting' size="20"></td>
    <td >

	</td>
  </tr>
  
  <tr>
	  <td width="100%" colspan=3 align="right" valign=absBottom>
		<BR><BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "AllowedNavigation()">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		
	</td>
  </tr>
</table>

</form>
</body>

</html>
