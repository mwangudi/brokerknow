<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Order Hold Option</title>
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
<script language="javascript">
	function  UpdateRequiresDate(theChk)
	{
		var holdVal = "0";
		if (theChk.checked)
		{
			holdVal = "1";
		}
		document.getElementById("RequiresDate").value = holdVal;
	}
	
	function  UpdateDefault(theChk)
	{
		var DefaultVal = "0"; 
		
		if (theChk.checked)
		{
			DefaultVal = "1";
		}
		
		document.getElementById("DefaultSelection").value = DefaultVal;
	}
</script>
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
                		window.self.close
                </script>
                <% response.end
        End If

	if action = "EXECUTE" then
		Dim description
        Dim RequiresDate
			 
		RequiresDate = cint(Request.Form("RequiresDate"))
        description = Request.Form("txtDescription")
        cmdCancel = Request.Form("cmdCancel")
        DefaultSelection = Trim(Request.Form("DefaultSelection"))
        
        If cmdCancel <> "" Then
			WriteDialogCancelScript
			Set Conn = Nothing
			Response.End
        End If
      
        'validate Description
        If Trim(Description) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Description"
                		window.self.close
                </script>
                <% response.end
        End If
        'validate size of Description
        If Len(Description) > 100 Then%>
                <script language = 'vbscript'>
                ShowMessage "Description can only be 100 characters in length"
                window.self.close
                </script>
                <% response.end
        End If
        
        if DefaultSelection = "" then DefaultSelection = 0
        
        'save data
         Set conn = GetActiveConnection("KBroker")
         
        sqlStr = "UPDATE [OrderHoldOptions] SET Description = " & "'" & description & "'" & _
				", RequiresDate = " & " " & RequiresDate & " " & _
				", DefaultSelection = " & " " & DefaultSelection & " " & _
				" WHERE OrderHoldOptionID  = " & ID                
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
               
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
   	end If
%>

<form name = 'frmEditOrderHoldOptions' id="frmMain" method = 'post' action = 'EditOrderHoldOptions.asp' >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<%
        Set conn = GetActiveConnection("KBroker")
       
        sqlStr = "SELECT OrderHoldOptionID, Description, RequiresDate, DefaultSelection FROM OrderHoldOptions WHERE OrderHoldOptionID = " & ID
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Order Hold Option cannot be retrieved for editing"
                		window.self.close
                </script>
                <% response.end
        End If
%>

  <tr>
    <td width="17%">Description</td>
    <td width="83%">&nbsp;&nbsp;<input type = 'text' name ='txtDescription' id = 'txtDescription' size="20" value = '<%=rs.Fields("Description")%>'></td>
  </tr>

  <tr>
    <td nowrap>Requires Date</td>
    <td>&nbsp;
     <%if cbool(Rs.Fields("RequiresDate")) then%>
		<input type=checkbox  checked value='True' name='chkRequiresDate' id='chkRequiresDate' onClick = 'UpdateRequiresDate(this);'> 
	<%else%>
		<input type=checkbox   value='False' name='chkRequiresDate' id='chkRequiresDate' onClick = 'UpdateRequiresDate(this);'> 
	<%end if%>
      </td>
  </tr>
  <br>
  <tr>
    <td nowrap>Default Selection</td>
    <td>&nbsp;&nbsp;
     
      <%if Rs.Fields("DefaultSelection") then%>
		<input type=checkbox  checked value='True' name='DefaultSel' id='DefaultSel' onClick = 'UpdateDefault(this);'> 
	<%else%>
		<input type='checkbox'   value='False' name='DefaultSel' id="DefaultSel" onClick = 'UpdateDefault(this);'>
	<%end if%>
      </td>
  </tr>
  
  <tr>
     <td width="100%" COLSPAN=2 align="right" valign=absBottom>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save">
		&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">		
		&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
		<input type = 'hidden' name ='RequiresDate' id = 'RequiresDate' value='<%= cint(Rs.Fields("RequiresDate").Value) %>'>
        <input type = 'hidden' name ='DefaultSelection' id = 'DefaultSelection' value='<%= cint(Rs.Fields("DefaultSelection").Value) %>'>
    </td>
  </tr>
</table>
</form>

</body>

</html>
