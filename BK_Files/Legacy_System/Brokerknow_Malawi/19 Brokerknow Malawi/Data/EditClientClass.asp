<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Client Class</title>
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
				
			document.frmEditClass.elements("CDAStatus").value = holdVal;
		}

function  UpdateAgentReturnable(theChk)
		{
			var holdVal = "0"; //order to be compounded
			if (theChk.checked)
			{
				holdVal = "1";//order not to be compounded
			}
				
			document.frmEditClass.elements("AgentStatus").value = holdVal;
		}

		function forceSubmit()
		{
			setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frmEditClass.method='post';
			document.frmEditClass.target='_self';
			document.frmEditClass.submit();		
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
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If

	if action = "EXECUTE" then
		Dim Voucher
       
		commission = Request.Form("cboCommission")
		description = Request.Form("txtDescription")
		Voucher =Request.Form("txtVoucher")

		toCancel = Request.Form("cmdCancel")
		CDA=Request.Form("CDAStatus")
       	AgentStatus =Request.Form("AgentStatus")     
        
                
		Set conn = GetActiveConnection("KBroker")
		If toCancel <> "" Then
			
			WriteDialogCancelScript			
			Response.End
		End If
       
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
                <% response.end
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
        
        sqlStr = "UPDATE [Class] SET ClassDescription = " & "'" & description & "'" & _
				",DefaultCommission = " & " " & commission & " " & _
				",IsCda = " & " " & CDA & " " & _
				",AgentStatus = " & " " & AgentStatus & " " & _
				",Voucher = " & " " & Voucher & " " & _
				" WHERE Class_DPA_  = " & ID                
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript2
        response.end
   	end If
   	
   
%>

<form name = 'frmEditClass' method = 'post' action = 'EditClientClass.asp' >
<table border="0" width="100%" cellspacing="1" cellpadding="1">
<%
        Set conn = GetActiveConnection("KBroker")
        
        
        sqlStr = "SELECT ClassDescription,Class_DPA_,DefaultCommission,IsCda,AgentStatus,Voucher FROM [Class] WHERE Class_DPA_  = " & ID
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Client class cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
        
       
%>
  <tr>
    <td width="17%"> Description</td>
    <td width="83%"><input type = 'text' name ='txtDescription' id = 'txtDescription' size="20" value = '<%=rs.Fields("ClassDescription")%>'></td>
  </tr>
  <tr>
    <td width="18%"> Default Commission</td>
    <td width="82%"><select name = 'cboCommission' tabIndex='3' id = 'cboCommission' size="1">
<%		Set conn = GetActiveConnection("KBroker")

        sqlStr = "SELECT * FROM [CommissionList]"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("Commission_DPA_") = rs.Fields("DefaultCommission") Then%>
                			<option selected value = '<%=rsEdit.Fields("Commission_DPA_")%>'><%=rsEdit.Fields("CommissionDisplay")%></option>
                		<%else%>
                        <option value = '<%=rsEdit.Fields("Commission_DPA_")%>'><%=rsEdit.Fields("CommissionDisplay")%></option>
                     <%end if
						rsEdit.MoveNext
                Loop
                if isnull(rs.Fields("DefaultCommission")) then%>
                	<option selected value = ''></option>
                <%end if
        End If
%>

    </select></td>
  </tr>
  <tr>
    <td>CDA</td>
    <% 
    Response.Write(rs("IsCda"))
    if cbool(rs("IsCda")) then %>
    <td><input type=checkbox   checked value='True' name='chkInterBank' onClick = 'UpdateCDA(this);'> 
      </td>
    <%
    else     
    %>
    <td><input type=checkbox   value='False' name='chkInterBank' onClick = 'UpdateCDA(this);'> 
      </td>
    <% end if%>
  </tr>    
  <tr>
    <td>Agent Returnable</td>
    <% if(rs("AgentStatus")=true) then %>
    <td><input type=checkbox   checked value='true' name='chkInterBank' onClick = 'UpdateAgentReturnable(this);'> 
      </td>
      <% else %>
      <td><input type=checkbox   value='False' name='chkInterBank' onClick = 'UpdateAgentReturnable(this);'> 
      </td>    
      <% end if%>
  </tr> 
  <tr>
    <td width="18%"> Voucher</td>
    <td width="82%"><input type = 'text' name ='txtVoucher' id = 'txtVoucher' size="20" value = '<%=rs.Fields("Voucher")%>'></td>
  </tr>     
  <tr>
    <td width="100%" colspan=2 align=right>
		<BR>
		<BR>
		<BR>
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save "  onclick="forceSubmit();">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">
    	<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
    	<input type = 'hidden' name ='CDAStatus' id = 'CDAStatus' value='0'>
		<input type = 'hidden' name ='AgentStatus' id = 'AgentStatus' value='0'>
    </td>
  </tr>
</table>
</form>

</body>

</html>
