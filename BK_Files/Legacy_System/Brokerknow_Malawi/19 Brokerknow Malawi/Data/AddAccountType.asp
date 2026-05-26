<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "AddAccountType"
	const DataEntity = "AccountType"
	const DataEntityPlural = "Account Types"
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
				Dim parentType
				Dim Setting
       
				parentType = Request.Form("cboParentType") 
				Setting = Request.Form("txtSetting")
		       
				'validate Parent Type
				If Trim(parentType) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please specify the Parent Type"
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
				
				sqlStr = "INSERT INTO [AccountType] (AccountTypeName,AccountTypeParent,AccountType_EIT_) " & _
											"		SELECT " & "'" & Setting & "'" & " as AccountTypeName " & _
											"		," & " " & parentType & " " & " as AccountTypeParent " & _
											"		," & "'" & guidStr & "'" & " as AccountType_EIT_" 
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
    <td >Parent Type</td>
    <td ><select name = 'cboParentType' id = 'cboParentType' size="1">
					<option selected value = ''></option>
<%
		Dim levelFilter
		
		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [AccountTypeList] WHERE AccountTypeLevel_DPA_ IN (1,2) Order By AccountTypeLevel_DPA_,AccountTypeSetting"
  
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                levelFilter = 1
                Do Until rs.EOF
				ShownAccount=rs.Fields("AccountTypeSetting") & ":" & rs.Fields("AccountTypeParent")

						if levelFilter <> rs.Fields("AccountTypeLevel_DPA_") then
								levelFilter = rs.Fields("AccountTypeLevel_DPA_")%>
								<option value = ''>**********</option>
						<%else%>
								<option value = '<%=rs.Fields("AccountType_DPA_")%>'><%=ShownAccount%></option>
                        <%rs.MoveNext
                        end if
                        
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
