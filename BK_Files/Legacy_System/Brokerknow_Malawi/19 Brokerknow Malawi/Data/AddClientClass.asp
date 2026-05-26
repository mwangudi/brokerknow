<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Client Class</title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
<script language="JavaScript" src="../Scripts/common.js"></script>

<Script>
function  UpdateCDA(theChk)
		{
			var holdVal = "0"; //order to be compounded
			if (theChk.checked)
			{
				holdVal = "1";//order not to be compounded
			}
				
			document.frmAddClass.elements("CDAStatus").value = holdVal;
		}

function  UpdateAgentReturnable(theChk)
		{
			var holdVal = "0"; //order to be compounded
			if (theChk.checked)
			{
				holdVal = "1";//order not to be compounded
			}
				
			document.frmAddClass.elements("AgentStatus").value = holdVal;
		}

		function forceSubmit()
		{
			setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frmAddClass.method='post';
			document.frmAddClass.target='_self';
			document.frmAddClass.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}
</Script>

</head>

<body Class="Dialog" onload="setOpener()">
<!--#include file="../libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim description
		Dim commission
		Dim Voucher
       
		commission = Request.Form("cboCommission")
		description = Request.Form("txtDescription")
		Voucher =Request.Form("txtVoucher")

        CDA=Request.Form("CDAStatus")
       	AgentStatus =Request.Form("AgentStatus")
        'validate Description
        If Trim(Description) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Description"
                		
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        'validate size of Description
        If Len(Description) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Description can only be 100 characters in length"
                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If      
	If trim(Voucher) <>""  and Not IsNumeric(Voucher) Then%>
		<script language = 'vbscript'>
		ShowMessage "Voucher can only be numeric"						
		</script>
		<% 
		ReloadPage(ID)
		response.end
	End If

	 Voucher =iif(trim(Voucher)="","Null",Voucher)

       commission = iif(commission = 0,"Null",commission)
       
        'save data
        sqlStr = "INSERT INTO [Class] (ClassDescription,Class_DPA_,DefaultCommission,IsCda,AgentStatus,Voucher) SELECT " & "'" & description & "'" & " as ClassDescription" & _
                "," & " " & "iif(isnull(max([Class_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Class'),max([Class_DPA_]) + 1)" & " " & " as Class_DPA_" & _
                "," & " " & commission & " " & " as DefaultCommission" & _
                "," & " " & CDA & " " & " as IsCda," & " " & AgentStatus & " " & " as AgentStatus," & " " & Voucher & " " & " as Voucher FROM [Class]"
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
                conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript2
   	end If
%>

<form name = 'frmAddClass' method = 'post' action = 'AddClientClass.asp' >

<table border="0" width="100%" cellspacing="1" cellpadding="1">
  <tr>
    <td width="18%"> Description</td>
    <td width="82%"><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20"></td>
  </tr>
  <tr>
    <td width="18%"> Default Commission</td>
    <td width="82%"><select name = 'cboCommission' id = 'cboCommission' size="1">
    	<option selected value = ''></option>
<%			Set conn = GetActiveConnection("KBroker")

        sqlStr = "SELECT * FROM [CommissionList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF
                        if cbool(rs.Fields("DefaultSelection")) then%>
								<option selected value = '<%=rs.Fields("Commission_DPA_")%>'><%=rs.Fields("CommissionDisplay")%></option>
                        <%else%>
								<option value = '<%=rs.Fields("Commission_DPA_")%>'><%=rs.Fields("CommissionDisplay")%></option>
                        <%end if
                        rs.MoveNext
                Loop
        End If
%>

    </select></td>
  </tr>
  <tr>
    <td>CDA</td>
    <td><input type=checkbox   value='False' name='chkInterBank' onClick = 'UpdateCDA(this);'> 
      </td>
  </tr>    
  <tr>
    <td>Agent Returnable</td>
    <td><input type=checkbox   value='False' name='chkInterBank' onClick = 'UpdateAgentReturnable(this);'> 
      </td>
  </tr>    
  <tr>
    <td width="18%"> Voucher</td>
    <td width="82%"><input type = 'text' name ='txtVoucher' id = 'txtVoucher' size="20"></td>
  </tr>
  <tr>
    <td width="100%" align=right colspan=2>
		<BR>
		<BR>
		<BR>
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save " onclick="forceSubmit();">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();" >
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='CDAStatus' id = 'CDAStatus' value='0'>
		<input type = 'hidden' name ='AgentStatus' id = 'AgentStatus' value='0'>
	</td>
    
  </tr>
</table>

</form>

</body>

</html>
