<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "AddAccount"
	const DataEntity = "Account"
	const DataEntityPlural = "Accounts"
	const ActionFolder = "Operations"
	
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim Level1
	Dim Level2
	Dim Level3
	Dim Code
	Dim AccName
	Dim OpeningBal
	Dim newLevel2
	Dim newLevel3
		
	action = ucase(Request.Form("action"))
	Level1 = Request.Form("cboAccountTypeLevel1")
	Level2 = Request.Form("cboAccountTypeLevel2")
	Level3 = Request.Form("cboAccountTypeLevel3")
	Code = Request.Form("txtCode")
	AccName = Request.Form("txtAccName")
	OpeningBal = Request.Form("txtOpeningBal")
	newLevel2 = Request.Form("txtAccountTypeLevel2")
	newLevel3 = Request.Form("txtAccountTypeLevel3")	
	if action = "EXECUTE" then
	
			'Response.Write Request.Form 
			'Response.End 
			
			Dim buttonAction
			Dim reloadRequired
		
			reloadRequired = false
			buttonAction = Trim(Ucase(Request.Form("buttonAction")))
			if buttonAction = "SAVE" then
					'validate Account Type
					If Trim(Level1) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please specify the Account Type"
					         		
					         </script>
					         <% ReloadPage(ID)
							 response.end
					End If
					
					If Trim(Level2) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please specify the Account Type level 2"
					         		
					         </script>
					         <% ReloadPage(ID)
							 response.end
					End If
					
					If Trim(Level3) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please specify the Account Type level 3"
					         		
					         </script>
					         <% ReloadPage(ID)
							 response.end
					End If
					
					'validate Code
					If Trim(Code) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please specify the Code"
					         		
					         </script>
					         <% ReloadPage(ID)
							 response.end
					End If
					'validate size of Code
					If Len(Code) > 50 Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Code can only be 50 characters in length"
					         		
					         </script>
					         <% ReloadPage(ID)
							 response.end
					End If
					
					'validate size of Level2
					If Len(newLevel2) > 100 Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Type2 can only be 100 characters in length"
					         		
					         </script>
					         <% ReloadPage(ID)
							 response.end
					End If
					'validate size of Level3
					If Len(newLevel3) > 100 Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Type3 can only be 100 characters in length"
					         		
					         </script>
					         <% ReloadPage(ID)
							 response.end
					End If
					'validate Name
					If Trim(AccName) = "" Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Please specify the Name"
					         		
					         </script>
					         <% ReloadPage(ID)
							 response.end
					End If
					'validate size of Name
					If Len(AccName) > 100 Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "Name can only be 100 characters in length"
					         		
					         </script>
					         <% ReloadPage(ID)
							 response.end
					End If
					
					'validate type of OpeningBal
					If Not IsNumeric(OpeningBal) Then%>
					         <script language = 'vbscript'>
					         		ShowMessage "The Opening Balance can only be numeric"
					         		
					         </script>
					         <% ReloadPage(ID)
							 response.end
					End If
					
					'save data	
					Dim guidStr 
					Dim guid 
					set guid = server.CreateObject("NDUtils.CGUID")
					
					Set conn = GetActiveConnection("KBroker")
					conn.BeginTrans
							if Level2 = "00" then
									If Trim(newLevel2) = "" Then%>
									         <script language = 'vbscript'>
									         		ShowMessage "Please specify Type 2"
									         		
									         </script>
									         <% ReloadPage(ID)
												response.end
									End If
									
									guidStr = guid.GenerateGUID
									sqlStr = "INSERT INTO [AccountType] (AccountTypeName,AccountTypeParent,AccountType_EIT_) " & _
											"		SELECT " & "'" & newLevel2 & "'" & " as AccountTypeName " & _
											"		," & " " & Level1 & " " & " as AccountTypeParent " & _
											"		," & "'" & guidStr & "'" & " as AccountType_EIT_" 
									
									
									sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
									conn.Execute sqlStr
									
									'obtain level2 key value
									sqlStr = "SELECT AccountType_DPA_ FROM AccountType WHERE AccountType_EIT_ = " & "'" & guidStr & "'"
									Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
									If (rs.EOF Or rs.BOF) Then%>
											<script language = 'vbscript'>
											        ShowMessage "A serious error has been encountered while saving the data. Try saving again"
												        
											</script>
											<% ReloadPage(ID)
												response.end
									End If
									Level2 = rs.Fields("AccountType_DPA_").value
							elseif Level2 = "" then
									Level2 = "Null"
							end if
							
							if Level3 = "00" then
									If Trim(newLevel3) = "" Then%>
									         <script language = 'vbscript'>
									         		ShowMessage "Please specify Type 3"
									         		
									         </script>
									         <% ReloadPage(ID)
											response.end
									End If
									
									if Level2 = "Null" then%>
									         <script language = 'vbscript'>
									         		ShowMessage "Please specify Type 2 first"
									         		
									         </script>
									         <% ReloadPage(ID)
											response.end
									else
											guidStr = guid.GenerateGUID
											sqlStr = "INSERT INTO [AccountType] (AccountTypeName,AccountTypeParent,AccountType_EIT_) " & _
													"		SELECT " & "'" & newLevel3 & "'" & " as AccountTypeName " & _
													"		," & " " & Level2 & " " & " as AccountTypeParent " & _
													"		," & "'" & guidStr & "'" & " as AccountType_EIT_" 
									
									
											sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
											conn.Execute sqlStr
									
											'obtain level3 key value
											sqlStr = "SELECT AccountType_DPA_ FROM AccountType WHERE AccountType_EIT_ = " & "'" & guidStr & "'"
											Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
											If (rs.EOF Or rs.BOF) Then%>
													<script language = 'vbscript'>
													        ShowMessage "A serious error has been encountered while saving the data. Try saving again"
														        
													</script>
													<% ReloadPage(ID)
														response.end
											End If
											Level3 = rs.Fields("AccountType_DPA_").value
									end if
							elseif Level3 = "" then
									Level3 = "Null"
							end if
														
							sqlStr = "INSERT INTO [Account] (AccountCode,AccountTypeLevel1,AccountTypeLevel2,AccountTypeLevel3,AccountName" & _
									"       ,Account_DPA_,AccountOpeningBal) SELECT " & "'" & Code & "'" & " as AccountCode," & " " & Level1 & " " & " as AccountTypeLevel1" & _
									"       ," & " " & Level2 & " " & " as AccountTypeLevel2" & _
									"       ," & " " & Level3 & " " & " as AccountTypeLevel3" & _
									"       ," & "'" & accName & "'" & " as AccountName," & " " & "iif(isnull(max([Account_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Account'),max([Account_DPA_]) + 1)" & " " & " as Account_DPA_" & _
									"       ," & " " & CDbl(OpeningBal) & " " & " as AccountOpeningBal" & _	
									"        FROM [Account]"
									
							
							sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
							conn.Execute sqlStr
					conn.CommitTrans
					conn.Close
					Set conn = Nothing
					WritefraEnabledDialogCloseScript2
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
			forceSubmit()
		}
		
		function FetchLevel()
		{
			document.frmMain.target = "_self"
			LevelSelected();
			
		}
		function forceSubmit()
		{
			//setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frm<%=DataSource%>.method='post';
			document.frm<%=DataSource%>.target='_self';
			document.frm<%=DataSource%>.submit();	
			
		}
		
		function setOpener()
		{

			window.self.opener = window.dialogArguments.opener;
					//alert(window.dialogArguments.opener.location);
		}
</script>
</head>

<body Class="Dialog"  onLoad="javascript: setOpener()" >

<form name = 'frm<%=DataSource%>' method = 'post' action = '<%=DataSource%>.asp' id = "frmMain">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
  
  <tr>
    <td width="15%">Name</td>
    <td width="54%"><input type = "text" STYLE="WIDTH: 400px;" name ="txtAccName" id = "txtAccName" size="20" value="<%=AccName%>"></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td width="15%">Code</td>
    <td width="54%"><input type = 'text' STYLE="WIDTH: 400px;" name ='txtCode' id = 'txtCode' size="20" value="<%=Code%>"></td>
    <td width="31%">

	</td>
  </tr>
   <tr>
    <td width="40%">Opening Balance</td>
    <%if OpeningBal = "" then 
			OpeningBal = 0
	end if%>
    <td width="54%"><input type = 'text' name ='txtOpeningBal' id = 'txtOpeningBal' size="20" value="<%=OpeningBal%>"></td>
    <td width="31%">
	</td>
  </tr>	
  <tr>
    <td width="15%">Type 1</td>
    <td width="54%"><select name = 'cboAccountTypeLevel1' id = 'cboAccountTypeLevel1' size="1" onchange='FetchLevel()'>
<%
		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [AccountTypeLevel1] Order By AccountTypeName"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
						if Level1 = "" then
								if cbool(rs.Fields("DefaultSelection")) then
										Level1 = rs.Fields("AccountType_DPA_")%>
										<option selected value = '<%=rs.Fields("AccountType_DPA_")%>'><%=rs.Fields("AccountTypeName")%></option>
								<%else%>
										<option value = '<%=rs.Fields("AccountType_DPA_")%>'><%=rs.Fields("AccountTypeName")%></option>
								<%end if
						else
								if rs.Fields("AccountType_DPA_").value = cint(Level1) then%>
										<option selected value = '<%=rs.Fields("AccountType_DPA_")%>'><%=rs.Fields("AccountTypeName")%></option>
								<%else%>
										<option value = '<%=rs.Fields("AccountType_DPA_")%>'><%=rs.Fields("AccountTypeName")%></option>
								<%end if
						end if
                        rs.MoveNext
                Loop
        End If
%>

    </select></td>
    <td width="31%">

	</td>
  </tr>
  <tr>
    <td width="15%">Type 2</td>
    <td width="54%"><select name = 'cboAccountTypeLevel2' id = 'cboAccountTypeLevel2' size="1" onchange='FetchLevel()'>
    <option selected value =''></option>
<%
        sqlStr = "SELECT * FROM [AccountTypeLevel2] WHERE AccountTypeParent = " &  Level1 & " Order By AccountTypeName"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                        if Level2 = "" then
								if cbool(rs.Fields("DefaultSelection")) then%>
										<option selected value = '<%=rs.Fields("AccountType_DPA_")%>'><%=rs.Fields("AccountTypeName")%></option>
								<%else%>
										<option value = '<%=rs.Fields("AccountType_DPA_")%>'><%=rs.Fields("AccountTypeName")%></option>
								<%end if
						else
								if rs.Fields("AccountType_DPA_").value = cint(Level2) then%>
										<option selected value = '<%=rs.Fields("AccountType_DPA_")%>'><%=rs.Fields("AccountTypeName")%></option>
								<%else%>
										<option value = '<%=rs.Fields("AccountType_DPA_")%>'><%=rs.Fields("AccountTypeName")%></option>
								<%end if
						end if
                        rs.MoveNext
                Loop
        End If
       
%>
	<%If Level2 = "00" then%>
			<option  selected value ='00'>Other</option>
	<%else%>
			<option  value ='00'>Other</option>
	<%end if%>
    </select></td>
    <td width="31%">
    <%If Level2 = "00" then%>
			<input type='text' size='20' name = 'txtAccountTypeLevel2' id = 'txtAccountTypeLevel2' style='display:all' value="<%=newLevel2%>">
	<%else%>
			<input type='text' size='20' name = 'txtAccountTypeLevel2' id = 'txtAccountTypeLevel2' style='display:none' value="<%=newLevel2%>">
	<%end if%>
	</td>
  </tr>
  <tr>
    <td width="15%">Type 3</td>
    <td width="54%"><select name = 'cboAccountTypeLevel3' id = 'cboAccountTypeLevel3' size="1" onchange='FetchLevel()'>
    <option selected value =''></option>
<%
		If Level2 = "" then
				Level2 = 0
		end if
		
        sqlStr = "SELECT * FROM [AccountTypeLevel3] WHERE AccountTypeParent = " &  Level2 & " Order By AccountTypeName"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                        if Level3 = "" then
								if cbool(rs.Fields("DefaultSelection")) then%>
										<option selected value = '<%=rs.Fields("AccountType_DPA_")%>'><%=rs.Fields("AccountTypeName")%></option>
								<%else%>
										<option value = '<%=rs.Fields("AccountType_DPA_")%>'><%=rs.Fields("AccountTypeName")%></option>
								<%end if
						else
								if rs.Fields("AccountType_DPA_").value = cint(Level3) then%>
										<option selected value = '<%=rs.Fields("AccountType_DPA_")%>'><%=rs.Fields("AccountTypeName")%></option>
								<%else%>
										<option value = '<%=rs.Fields("AccountType_DPA_")%>'><%=rs.Fields("AccountTypeName")%></option>
								<%end if
						end if
                        rs.MoveNext
                Loop
        End If
       
%>
	<%If Level3 = "00" then%>
			<option  selected value ='00'>Other</option>
	<%else%>
			<option  value ='00'>Other</option>
	<%end if%>
    </select></td>
    <td width="31%">
    <%If Level3 = "00" then%>
			<input type='text' size='20' name = 'txtAccountTypeLevel3' id = 'txtAccountTypeLevel3' style='display:all' value="<%=newLevel3%>">
	<%else%>
			<input type='text' size='20' name = 'txtAccountTypeLevel3' id = 'txtAccountTypeLevel3' style='display:none' value="<%=newLevel3%>">
	<%end if%>
	</td>
  </tr>  
  <tr>
	  <td width="100%" colspan=3 align="right" valign=absBottom>
		<BR><BR>
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "AllowedNavigation()">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='buttonAction' id = 'action' value="Save">
	</td>
  </tr>
</table>

</form>
</body>

</html>
