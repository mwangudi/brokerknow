<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Order Hold Options</title>

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
	
	action = trim(ucase(Request.Form("action")))
	
	if action = "EXECUTE" then
		Dim Description
		Dim RequiresDate
		Dim DefaultSelection
			 
		Description = Request.Form("txtDescription")
		RequiresDate = Request.Form("RequiresDate")
		DefaultSelection = Trim(Request.Form("DefaultSelection"))
               
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
         conn.BeginTrans
         
        'if cint(DefaultSelection) = 1 then 
         'There can only be one default selection record in the table orderhold options
          
       '   sqlstr = "Update OrderHoldOptions set DefaultSelection = 0 "
          
        '  conn.Execute sqlstr
          
       ' end if 
         
        sqlStr = "INSERT INTO [OrderHoldOptions] (Description,RequiresDate,OrderHoldOptionID,DefaultSelection) SELECT " & "'" & Description & "'" & " as Description" & _
				"," & " " & RequiresDate & " " & " as RequiresDate" & _
                "," & " " & "iif(isnull(max([OrderHoldOptionID])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'OrderHoldOptions'),max([OrderHoldOptionID]) + 1)" & " " & " as OrderHoldOptionID" & _
                "," & " " & DefaultSelection & " " & "  as DefaultSelection FROM [OrderHoldOptions]"
        
        conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
        conn.CommitTrans
        conn.Close
        
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End
   	end If
%>
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
<form name = 'frmAddOrderHoldOptions' method = 'post' action = 'AddOrderHoldOptions.asp' >
<BR>
<table border="0" width="100%" cellspacing="2" cellpadding="0">
  <tr>
    <td width="18%">Description</td>
    <td width="82%">&nbsp;&nbsp;<input type = 'text' name ='txtDescription' id = 'txtDescription' size="20"></td>
  </tr>
  <tr>
    <td nowrap>Requires Date</td>
    <td>&nbsp;&nbsp;<input type=checkbox   value='False' name='chkRequiresDate' onClick = 'UpdateRequiresDate(this);'> 
      </td>
  </tr>
  <tr>
    <td nowrap>Default Selection</td>
    <td>&nbsp;&nbsp;<input type='checkbox'   value='False' name='DefaultSel' id="DefaultSel" onClick = 'UpdateDefault(this);'> 
      </td>
  </tr>
   <tr>
    <td width="100%" COLSPAN=2 align="right" valign=absBottom>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save">
		&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Close " OnClick="JavaScript: window.close()">		
		&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='RequiresDate' id = 'RequiresDate' value='0'>
		<input type = 'hidden' name ='DefaultSelection' id = 'DefaultSelection' value='0'>
    </td>
  </tr>
  </table>
</form>
</body>

</html>
