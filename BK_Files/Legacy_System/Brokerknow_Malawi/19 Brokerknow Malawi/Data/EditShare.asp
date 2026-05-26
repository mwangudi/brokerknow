<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Announcement</title>

  <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
 
  <!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">


</head>

<body Class="Dialog">
<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
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
                <% 
				WriteDialogRefuseOpenScript
				response.end
        End If

	if action = "EXECUTE" then
		Dim security
		Dim announceDate
		Dim closing
		Dim pay
		Dim yend
		Dim announcement
		
        
       announcement = Request.Form("cboAnnouncement")        
       security = Request.Form("cboSecurity") 
       announceDate = Request.Form("txtAnnouncement")
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
        
        Set conn = GetActiveConnection("KBroker")
       
        'save data
        sqlStr = "UPDATE [Share] SET ShareAnnouncement = " & "#" & FormatDate(announceDate) & "#" & ",ShareClosing = " & "#" & FormatDate(closing) & "#" & "" & _
                "       ,SharePDate = " & "#" & FormatDate(pay) & "#" & ",ShareYEnd = " & "#" & yend & "#" & "" & _
                "       ,Security_DPA_ = " & " " & security & " " & ",ShareAnnouncementType_DPA_ = " & " " & announcement & " " & " WHERE Share_DPA_  = " & ID       
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End
   	end If
%>


<form name = 'frmEditSecAnnounce' method = 'post' action = 'EditShare.asp' >
<table border="0" width="100%" cellspacing="1" cellpadding="0">
  <tr>
    <td width="17%">Security</td>
    <td width="83%"><select name = 'cboSecurity' id = 'cboSecurity' size="1">
<%
        Set conn = GetActiveConnection("KBroker")
       
        
        sqlStr = "SELECT ShareAnnouncement,ShareClosing,SharePDate,ShareYEnd,Share_DPA_" & _
                "       ,SecurityList.Security_DPA_,ShareAnnouncementType_DPA_ FROM [SecurityList] INNER JOIN [Share] ON SecurityList.Security_DPA_ = Share.Security_DPA_ WHERE Share_DPA_  = " & ID
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Announcement cannot be retrieved for editing"
                		
                </script>
                <% WriteDialogRefuseOpenScript
                response.end
        End If
        
       
        
        sqlStr = "SELECT * FROM [SecurityList] WHERE OrderSecTypeDescription = 'S'"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("Security_DPA_") = rs.Fields("Security_DPA_") Then%>
                			<option selected value = '<%=rsEdit.Fields("Security_DPA_")%>'><%=rsEdit.Fields("SecurityName")%></option>
                		<%else%>
                        <option value = '<%=rsEdit.Fields("Security_DPA_")%>'><%=rsEdit.Fields("SecurityName")%></option>
                     <%end if
						rsEdit.MoveNext
                Loop
        End If
%>

    </select></td>
  </tr>
  <tr>
    <td width="17%">Announcement</td>
    <td width="83%"><select name = 'cboAnnouncement' id = 'cboAnnouncement' size="1">
<%
        sqlStr = "SELECT * FROM [ShareAnnouncementTypeList]"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("ShareAnnouncementType_DPA_") = rs.Fields("ShareAnnouncementType_DPA_") Then%>
                			<option selected value = '<%=rsEdit.Fields("ShareAnnouncementType_DPA_")%>'><%=rsEdit.Fields("ShareAnnouncementTypeName")%></option>
                		<%else%>
                        <option value = '<%=rsEdit.Fields("ShareAnnouncementType_DPA_")%>'><%=rsEdit.Fields("ShareAnnouncementTypeName")%></option>
                     <%end if
						rsEdit.MoveNext
                Loop
        End If
%>

    </select></td>
  </tr>
<script language="JavaScript" src="../CALENDAR/calendar.js"></script>
<SCRIPT language="JavaScript">
	var calAnnouncement=new ctlSpiffyCalendarBox("calAnnouncement", "frmEditSecAnnounce", "txtAnnouncementDate","cmdAnnouncement","<%= FormatDate(rs.Fields("ShareAnnouncement")) %>",1);
	var calClosing=new ctlSpiffyCalendarBox("calClosing", "frmEditSecAnnounce", "txtClosing","cmdClosing","<%= FormatDate(rs.Fields("ShareClosing")) %>",1);
	var calPDate=new ctlSpiffyCalendarBox("calPDate", "frmEditSecAnnounce", "txtPDate","cmdPDate","<%= FormatDate(rs.Fields("SharePDate")) %>",1);
	var calYEnd=new ctlSpiffyCalendarBox("calYEnd", "frmEditSecAnnounce", "txtYEnd","cmdYEnd","<%= FormatDate(rs.Fields("ShareYEnd")) %>",1);
</SCRIPT>
<!--END CALENDAR -->
  <tr>
    <td width="17%">Announcement Date</td>
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
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
      </td>
  </tr>
</table>
</form>

</body>

</html>














