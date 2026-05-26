<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Announcement</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>

<!--END CALENDAR -->
</head>

<body Class="Dialog">
<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim security
		Dim announceDate
		Dim closing
		Dim pay
		Dim yend
		Dim announcement
		
        
       announcement = Request.Form("cboAnnouncement")
       security = Request.Form("cboSecurity") 
       announceDate = Request.Form("txtAnnouncementDate")
       closing = Request.Form("txtClosing")
       pay = Request.Form("txtPDate")
       yend = Request.Form("txtYEnd")
		
       
        'validate Announcement
        If Trim(Announcement) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Announcement"
                </script>
                <% response.end
        End If
       'validate Security
        If Trim(Security) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Security"
                </script>
                <% response.end
        End If
               
        'save data
        sqlStr = "INSERT INTO [Share] (ShareAnnouncement,ShareClosing,SharePDate,ShareYEnd,Share_DPA_" & _
                "       ,Security_DPA_,ShareAnnouncementType_DPA_) SELECT " & "#" & FormatDate(announceDate) & "#" & " as ShareAnnouncement" & _
                "       ," & "#" & FormatDate(closing) & "#" & " as ShareClosing" & _
                "       ," & "#" & FormatDate(pay) & "#" & " as SharePDate," & "#" & yend & "#" & " as ShareYEnd" & _
                "       ," & " " & "iif(isnull(max([Share_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Share'),max([Share_DPA_]) + 1)" & " " & " as Share_DPA_" & _
                "       ," & " " & security & " " & " as Security_DPA_" & _
                "       ," & " " & announcement & " " & " as ShareAnnouncementType_DPA_" & _
                "        FROM [Share]"
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
                conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End
   	end If
%>

<SCRIPT language="JavaScript">
	var calAnnouncement=new ctlSpiffyCalendarBox("calAnnouncement", "frmAddAnnouncement", "txtAnnouncementDate","cmdAnnouncement","<%= FormatDate(Date) %>",1);
	var calClosing=new ctlSpiffyCalendarBox("calClosing", "frmAddAnnouncement", "txtClosing","cmdClosing","<%= FormatDate(Date) %>",1);
	var calPDate=new ctlSpiffyCalendarBox("calPDate", "frmAddAnnouncement", "txtPDate","cmdPDate","<%= FormatDate(Date) %>",1);
	var calYEnd=new ctlSpiffyCalendarBox("calYEnd", "frmAddAnnouncement", "txtYEnd","cmdYEnd","<%= FormatDate(Date) %>",1);
</SCRIPT>

<form name = 'frmAddAnnouncement' method = 'post' action = 'AddShare.asp' >
<table border="0" width="100%" cellspacing="1" cellpadding="0">
  <tr>
    <td width="17%">Security</td>
    <td width="83%"><select name = 'cboSecurity' id = 'cboSecurity' size="1">
    	<option selected value = ''></option>
<%
        Set conn = GetActiveConnection("KBroker")
        
        sqlStr = "SELECT * FROM [SecurityList] WHERE OrderSecTypeDescription = 'S'"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF%>
                        <option value = '<%=rs.Fields("Security_DPA_")%>'><%=rs.Fields("SecurityName")%></option>
                        <%rs.MoveNext
                Loop
        End If
%>

    </select></td>
  </tr>
  <tr>
    <td width="17%">Announcement</td>
    <td width="83%"><select name = 'cboAnnouncement' id = 'cboAnnouncement' size="1">
    	<option selected value = ''></option>
<%
        sqlStr = "SELECT * FROM [ShareAnnouncementTypeList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF%>
                        <option value = '<%=rs.Fields("ShareAnnouncementType_DPA_")%>'><%=rs.Fields("ShareAnnouncementTypeName")%></option>
                        <%rs.MoveNext
                Loop
        End If
%>

    </select></td>
  </tr>
  <tr>
    <td width="17%%">Announcement Date</td>
    <td width="83%"><SCRIPT language="JavaScript">calAnnouncement.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td width="17%">Closing</td>
    <td width="83%"><SCRIPT language="JavaScript">calClosing.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td width="17%">Payment&nbsp;</td>
    <td width="83%"><SCRIPT language="JavaScript">calPDate.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td width="17%">Year End</td>
    <td width="83%"><SCRIPT language="JavaScript">calYEnd.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td width="100%" colspan="2" align=right>
		<BR>
		<BR>
		<BR>
		<BR>
		<BR>
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

</html>
