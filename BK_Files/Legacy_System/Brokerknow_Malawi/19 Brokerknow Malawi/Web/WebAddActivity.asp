<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Activity</title>


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
	
   Dim action
   Dim conn 
   Dim sqlStr
   Dim rs
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim client
		Dim name
		Dim actDate
		Dim notes 
        
		client = Request.Form("cboClient")
        name = Request.Form("cboActvtyClass")
        actDate = Request.Form("txtDate")
        notes = Request.Form("txtNotes")
        
		'validate Client		
        If Trim(client) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Client"
                		window.history.go(-1)
                </script>
                <% response.end
        End If
        'validate Name
        If Trim(name) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Name"
                		
                </script>
                <% response.end
        End If
        'validate size of Notes
        If Len(Notes) > 255 Then%>
                <script language = 'vbscript'>
                ShowMessage "Notes can only be 255 characters in length"
                
                </script>
                <% response.end
        End If

        'save data
        sqlStr = "INSERT INTO [Activity] (ActivityDate,ActivityNotes,Activity_DPA_,ActvtyClass_DPA_,Client_DPA_) SELECT " & "#" & actDate & "#" & " as ActivityDate," & "'" & Notes & "'" & " as ActivityNotes" & _
                "       ," & " " & "iif(isnull(max([Activity_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Activity'),max([Activity_DPA_]) + 1)" & " " & " as Activity_DPA_" & _
                "       ," & " " & name & " " & " as ActvtyClass_DPA_" & _
                "       ," & " " & client & " " & " as Client_DPA_" & _
                "        FROM [Activity]"
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
                conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
       ' WritefraEnabledDialogCloseScript
        Response.redirect "../WebMenu.asp"
	end If%>

<form name = 'frmAddActivity' method = 'post' action = 'WebAddActivity.asp'>
<table border="0" width="100%">
  <tr>
    <td width="17%">Client</td>
    <td width="83%">
<%
        Set conn = GetActiveConnection("KBroker")
        
        sqlStr = "SELECT * FROM [FullClientList] WHERE Client_DPA_ = " & Session("Client_DPA_")
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		alert("There was an error accessing your account.")
                
                </script>
                <% response.end
                        
        End If
%>
<select name = 'cboClient' id = 'cboClient' style="display:none">
<option value = '<%=rs.Fields("Client_DPA_")%>'><%=rs.Fields("ClientName")%></option>
    </select>
    <%=rs.Fields("ClientName")%>
    </td>
  </tr>
 <tr>
    <td width="17%">Type</td>
    <td width="83%"><select name = 'cboActvtyClass' id = 'cboActvtyClass'>
    	<option selected value = ''></option>
<%
    
        sqlStr = "SELECT * FROM [ActvtyClassList] WHERE ClientAccess = 1"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF%>
                        <option value = '<%=rs.Fields("ActvtyClass_DPA_")%>'><%=rs.Fields("ActvtyClassDescription")%></option>
                        <%rs.MoveNext
                Loop
        End If
        conn.Close
        Set conn = Nothing
%>

    </select></td>
  </tr>
    <tr>
    <td width="17%">Date</td>
    <td width="83%"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
  </tr>
 
  <tr>
    <td width="17%">Notes</td>
    <td width="83%"><textarea name ='txtNotes' id = 'txtNotes' rows="5" cols="34"></textarea></td>
  </tr>
  <tr>
    <td width="100%" COLSPAN=2 align="right" valign=absBottom>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save">
		&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.location.replace('../webmenu.asp')">		
		&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
    </td>
  </tr>
</table>
</form>

</body>

</html>
