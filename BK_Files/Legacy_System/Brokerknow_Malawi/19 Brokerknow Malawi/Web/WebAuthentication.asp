<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Web Authentication</title>


<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmAddActivity", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
</SCRIPT>
<!--END CALENDAR -->
</head>

<body >

<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<%
	
	Session("UserID") = -1
	
   Dim action
   Dim conn 
   Dim sqlStr
   Dim rs
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then		
        
		clientCDS = Request.Form("txtCdsNo")
        ClientID = Request.Form("txtCertNo")
        
        Set conn = GetActiveConnection("KBroker")
        
		'validate Client		
        If Trim(clientCDS) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please Type in the REF NO"
                		window.history.go(-1)
                </script>
                <% response.end
        End If
        'validate Name
        If Trim(ClientID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the REF ID No"
                		window.history.go(-1)
                </script>
                <% response.end
        End If
        
        sqlStr="SELECT Client_DPA_, ClientCDSNo, ClientIDPass FROM Client " & _
			   " WHERE (ClientCDSNo LIKE N'%" & clientCDS & "%') AND (ClientIDPass = N'" & ClientID & "')"        
        
        'Response.Write(sqlStr)
        'Response.End
        
        Set rs = conn.Execute(sqlStr)                        
        
        if not (rs.eof and rs.bof) then
        Response.redirect "../Web/AddOnlineUser.asp?Clientid=" & rs("Client_DPA_")
        else
		%>
                <script language = 'vbscript'>
                		ShowMessage "The details don't match any in the database"
                		window.history.go(-1)
                </script>
                <% response.end        
        end if
        
       ' WritefraEnabledDialogCloseScript
        
	end If%>

<form name = 'frmWebAuthentication' method = 'post' action = 'WebAuthentication.asp'>
<table border="0" width="100%">
<br>
<br>
<br>
<br>
<br>
<br>
<tr>
<td align="center">
<table border="0" width="300">
<tr><td colspan="2" align="center"><b>Online Vefication</b></td></tr>
<tr><td colspan="2">&nbsp;</td></tr>
<tr>
		<td width="20%">REF NO</td>
		<td width="30%"><input type = 'text' name ='txtCdsNo' STYLE="WIDTH: 200px" tabIndex='8' id = 'txtCdsNo'></td>
	</tr>
	<tr>
		<td width="20%">ID NO (Certificate)</td>
		<td width="30%"><input type = 'text' name ='txtCertNo' STYLE="WIDTH: 200px" tabIndex='9' id = 'txtCertNo' size="20"></td>
	</tr>	
  <tr>
    <td width="100%" COLSPAN=2 align="right" valign=absBottom>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="  OK  ">
		&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.location.replace('../webdefault.htm')">		
		&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
    </td>
  </tr>
  </td>
  </tr>
  </table>
</table>
</form>

</body>

</html>
