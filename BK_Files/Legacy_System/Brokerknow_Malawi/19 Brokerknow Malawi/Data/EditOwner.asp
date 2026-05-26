<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Owner</title>
   <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
</head>

<body Class="Dialog">

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
		Dim fname
       Dim lname
       Dim rate
       Dim OpeningBal
	   Dim MobileNo
	   Dim SendSMS
       
       OpeningBal = Request.Form("txtOpeningBal")
       rate = Request.Form("txtRate")
       fname = Request.Form("txtFName")
       lname = Request.Form("txtLName")
       IdNo = Request.Form("txtId")
	   MobileNo = trim(replace(Request.Form("txtMobileNo")," ",""))
	   SendSMS = Request.Form("cboSendSMS")

       'validate Rate
        If Trim(Rate) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Rate"
                		
                </script>
                <% response.end
        End If
        'validate First Name
        If Trim(fname) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the First Name"
                		
                </script>
                <% response.end
        End If
        
        'validate Last Name
        If Trim(lname) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Last Name"
                		
                </script>
                <% response.end
        End If
        'validate size of First Name
        If Len(Fname) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "First Name can only be 100 characters in length"
                
                </script>
                <% response.end
        End If
         'validate size of Last Name
        If Len(lName) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Last Name can only be 100 characters in length"
                
                </script>
                <% response.end
        End If
        
        If Not IsNumeric(OpeningBal) Then%>
				<script language = 'vbscript'>
				ShowMessage "Opening Balance can only be numeric"						
				</script>
				<% response.end
		End If
		
        Set conn = GetActiveConnection("KBroker")
       
        'save data
        
        sqlStr = "UPDATE [Owner] SET OwnerFname = " & "'" & fname & "'" & _
				",OwnerLName = " & "'" & lname & "'" & "" & _
				",IdNo = " & "'" & IdNo & "'" & "" & _
				",CommissionRate = " & " " & rate & " " & "" & _
				",OwnerOpeningBal = " & " " & OpeningBal & " " & "" & _
				",MobileNo = " & " " & MobileNo & " " & "" & _
				",SendSMS = " & " " & SendSMS & " " & "" & _
                " WHERE Owner_DPA_  = " & ID
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing
        WriteFraEnabledDialogCloseScript
        response.end
   	end If
%>



<form name = 'frmEditOwner' method = 'post' action = 'EditOwner.asp' target="deleteFrame" OnSubmit="JavaScript: UpdateDialogHandle();">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
<%
        Set conn = GetActiveConnection("KBroker")
       
        sqlStr = "SELECT OwnerFname,OwnerLName,Owner_DPA_,CommissionRate,OwnerOpeningBal,IdNo,MobileNo,SendSMS FROM [Owner] WHERE Owner_DPA_  = " & ID        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Owner cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If

		
%>
  <tr>
    <td width="30%"> First Name</td>
    <td width="80%"><input type = 'text' name ='txtFname' id = 'txtFname' size="20" value = '<%=rs.Fields("OwnerFname")%>'></td>
  </tr>
  <tr>
    <td width="30%"> Last Name</td>
    <td width="80%"><input type = 'text' name ='txtLName' id = 'txtLName' size="20" value = '<%=rs.Fields("OwnerLName")%>'></td>
  </tr>
  <tr>
    <td width="30%" >ID No</td>
    <td width="70%" ><input type = 'text' name ='txtId' id = 'txtId' size="20" value ='<%=rs.Fields("IdNo")%>'></td>
  </tr>  
  <tr>
    <td width="18%">Commission Rate</td>
    <td width="82%"><input type = 'text' name ='txtRate' id = 'txtRate' STYLE="TEXT-ALIGN: RIGHT;" size="20" value="<%=rs.Fields("CommissionRate")%>"></td>
  </tr>
  <tr>
    <td nowrap> Opening Balance</td>
    <td nowrap><input type = 'text' name ='txtOpeningBal' STYLE="TEXT-ALIGN: RIGHT;" id = "txtOpeningBal" size="20" value="<%=rs.Fields("OwnerOpeningBal")%>"></td>
  </tr>
  <tr>
    <td nowrap> Mobile Number</td>
    <td nowrap><input type = 'text' name ='txtMobileNo' STYLE="TEXT-ALIGN: RIGHT;" id = "txtMobileNo" size="20" value="<%=rs.Fields("MobileNo")%>"></td>
  </tr>
  <tr>
    <td nowrap> Send SMS</td>
    <td nowrap><select name="cboSendSMS" id="cboSendSMS">
	<%
	 if trim(rs.fields("SendSMS")) = 1 then
	   YesSMS = "Selected"
	 else
	   NoSMS = "Selected"
	 end if
	%>
	<option value="0"  <%=NoSMS%>>No</option>
	<option value="1" <%=YesSMS%>>Yes</option>
	</select></td>
  </tr>
  <tr>
    <td width="100%" colspan="2" align=right>
		<BR>
		<BR>		
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
      </td>
  </tr>
</table>
</form>

</body>

</html>
