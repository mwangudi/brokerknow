<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Owner</title>
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
	
	action = ucase(Request.Form("action"))
	
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
        If Len(lname) > 100 Then%>
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
       
        'save data
       sqlStr = "INSERT INTO [Owner] (OwnerFname,OwnerLName,IdNo,CommissionRate,OwnerOpeningBal,Owner_DPA_,MobileNo,SendSMS) SELECT " & _
				"'" & fname & "'" & " as OwnerFname" & _
				"," & "'" & lname & "'" & " as OwnerLName" & _
				"," & "'" & IdNo & "'" & " as IdNo" & _
				"," & " " & rate & " " & " as CommissionRate" & _
				"," & " " & OpeningBal & " " & " as OwnerOpeningBal" & _
                "," & " " & "iif(isnull(max([Owner_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Owner'),max([Owner_DPA_]) + 1)" & " " & " as Owner_DPA_" & _
				"," & " " & MobileNo &  " " & " as MobileNo " & _
				"," & " " & SendSMS &  " " & " as SendSMS " & _
                " FROM [Owner]"
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
                conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        WriteFraEnabledDialogCloseScript
        Response.End
   	end If
%>

<form name = 'frmAddOwner' method = 'post' action = 'AddOwner.asp' target="deleteFrame" OnSubmit="JavaScript: UpdateDialogHandle();">
<table border="0" width="100%" cellspacing="1" cellpadding="1">
  <tr>
    <td width="30%" > First Name</td>
    <td width="70%" ><input type = 'text' name ='txtFname' id = 'txtFname' size="20"></td>
  </tr>
  <tr>
    <td width="30%" > Last Name</td>
    <td width="70%" ><input type = 'text' name ='txtLName' id = 'txtLName' size="20"></td>
  </tr>
  <tr>
    <td width="30%" >ID No</td>
    <td width="70%" ><input type = 'text' name ='txtId' id = 'txtId' size="20"></td>
  </tr>  
  <tr>
    <td width="18%">Commission Rate</td>
    <td width="82%"><input type = 'text' name ='txtRate' id = 'txtRate' STYLE="TEXT-ALIGN: RIGHT;" size="20" value="0"></td>
  </tr>
  <tr>
    <td nowrap> Opening Balance</td>
    <td nowrap><input type = 'text' name ='txtOpeningBal' STYLE="TEXT-ALIGN: RIGHT;" id = "txtOpeningBal" size="20" value="0"></td>
  </tr>
  <tr>
    <td nowrap> Mobile Number</td>
    <td nowrap><input type = 'text' name ='txtMobileNo' STYLE="TEXT-ALIGN: RIGHT;" id = "txtMobileNo" size="20" value="256"></td>
  </tr>
  <tr>
    <td nowrap> Send SMS</td>
    <td nowrap><select name="cboSendSMS" id="cboSendSMS">
	<option value="0" selected>No</option>
	<option value="1" >Yes</option>
	</select></td>
  </tr>
  <tr>
   <td width="100%" colspan="2" align=right>
		<BR>
	
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
      </td>
  </tr>

</table>
</form>

</body>

