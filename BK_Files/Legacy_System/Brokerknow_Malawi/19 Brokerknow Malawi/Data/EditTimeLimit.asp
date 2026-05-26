<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Edit Time Limit</title>

  <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
<script language='javascript'>
	function forceSubmit()
		{
			setOpener();
			document.frmEditTimeLimit.method='post';
			document.frmEditTimeLimit.target='_self';
			document.frmEditTimeLimit.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}
</script>

</head>

<body Class="Dialog"  onload="setOpener()">
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
                <% WriteDialogRefuseOpenScript
                response.end
        End If

	if action = "EXECUTE" then
		Dim security
		Dim number
		Dim idate
		Dim life
		Dim pay
		Dim rate
		
        
       intLimit = Request.Form("txtIntLimit") 
       nseLimit = Request.Form("txtNSELimit")
             
       'validate
        If Not IsNumeric(intLimit) Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify a valid Internal Limit"
                
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
   
        'validate 
        If Not IsNumeric(nseLimit) Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify a valid NSE Limit"
                 
                </script>
                <% 
				reloadPage(ID)
				response.end
        End If
  
        
        Set conn = GetActiveConnection("KBroker")
        
        'save data
       sqlStr = "UPDATE [TimeLimit] SET TimeLimitLimDaysInt = "  & intLimit &  ", TimeLimitLimDaysNSE = " & nseLimit  & "" & _
				" WHERE TimeLimit_DPA_  = " & ID
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript2
        Response.End
   	end If
   	
   	 Set conn = GetActiveConnection("KBroker")
   	
   	 sqlStr = "SELECT * FROM [TimeLimitList] WHERE TimeLimit_DPA_ = " & ID
     Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
     
     If rsEdit.EOF Or rsEdit.BOF Then    
     %>	
		<Script Language="JavaScript">
			ShowMessage("The Time Limit cannot be retrieved for editing.")
		</Script>
      
      <%Set rsEdit = Nothing
		Set Conn = Nothing				
		WriteDialogRefuseOpenScript
		Response.End
     End If   
%>

<form name = 'frmEditTimeLimit' method = 'post' action = 'EditTimeLimit.asp' >
<table border="0" width="100%">
	<tr>
		<td width="30%">Action</td>
		<td width="70%"><input type = 'text' class="Readonly" READONLY name ='txtAction' id = 'txtAction' size="20" value="<%= rsEdit.Fields("TimeLimitAction").Value %>"></td>
	</tr>	
  <tr>
    <td width="35%">Internal Limit (days)</td>
    <td width="70%"><input type = 'text' name ='txtIntLimit' id = 'txtIntLimit' size="20" value="<%= rsEdit.Fields("TimeLimitInternal").Value %>"></td>
  </tr>
  
   <tr>
    <td width="30%">NSE Limit (days)</td>
    <td width="70%"><input type = 'text' name ='txtNSELimit' id = 'txtNSELimit' size="20" value="<%= rsEdit.Fields("TimeLimitNSE").Value %>"></td>
  </tr>

 </table>


<table border=0 width="100%">  
  <tr>
     <td align=right>
		<BR>
		<BR>
		<BR>
		<BR>
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick = "forceSubmit()">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">		
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
		<input type = 'hidden' name ='buttonAction' id = 'buttonAction' value="Save">
     </td>
  </tr>
</table>
</form>
</body>

</html>














