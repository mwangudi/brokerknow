<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Activity</title>
  
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>

</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->

<div id="spiffycalendar" class="text">&nbsp;</div>

<%
			
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	UserId=Session("UserID")	
	
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
		Dim client
		Dim name
		Dim actDate
		Dim notes 
        
        client = Request.Form("cboClient")
        name = Request.Form("cboActvtyClass")
        actDate = Request.Form("txtDate")
        notes = Request.Form("txtNotes")
        toCancel = Request.Form("cmdCancel")
        status=Request.Form("cboStatus")
        
        Set conn = GetActiveConnection("KBroker")
        
        If toCancel <> "" Then
			
			WriteDialogCancelScript
			Response.End
        End If
        		
        'validate Name
        If Trim(name) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Name"
                		window.self.close
                </script>
                <% response.end
        End If
        'validate size of Notes
        If Len(Notes) > 255 Then%>
                <script language = 'vbscript'>
                ShowMessage "Notes can only be 255 characters in length"
                window.self.close
                </script>
                <% response.end
        End If
        
        

        'save data
        sqlStr = "UPDATE Activity SET ActivityDate = " & "#" & actDate & "#" & ",ActivityNotes = " & "'" & Notes & "'" & "" & _
                "       ,ActvtyClass_DPA_ = " & " " & name & " " & "," & _
                "        status=" & " " & status & " " & ",Changedby=" & " " & UserID & " " & ",TimeModified=GetDate() WHERE Activity_DPA_  = " & ID
         
		
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))               
                
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
	end If
	%>

<form name = 'frmEditActivity' method = 'post' action = 'EditActivity.asp' >
<table border="0" width="100%">
  <tr>
    <td width="17%">Client</td>
    
<%
        Set conn = GetActiveConnection("KBroker")
        
        Dim timeNow
        timeNow = Timer
       
  
		sqlStr = "SELECT ActivityDate,ActivityNotes,Activity_DPA_,Activity.ActvtyClass_DPA_,Activity.status, Client.Client_DPA_,Client.ClientName FROM [Client] INNER JOIN [ActvtyClassList] INNER JOIN [Activity] ON ActvtyClassList.ActvtyClass_DPA_ = Activity.ActvtyClass_DPA_ ON Client.Client_DPA_ = Activity.Client_DPA_ WHERE Activity_DPA_  = " & ID

        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Activity cannot be retrieved for editing"
                		window.self.close
                </script>
                <% response.end
        End If       
     %>  
  
  <td>	
      <input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text' name ='txtClient' id = 'txtClient' size="9" value="<%=rs("ClientName")%>">
      </td>
    	    
  </tr>


<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmEditActivity", "txtDate","cmdDate","<%= FormatDate(rs.Fields("ActivityDate")) %>",1);
</SCRIPT>
<!--END CALENDAR -->

 <tr>
    <td width="17%">Type</td>
    <td width="83%"><select name = 'cboActvtyClass' id = 'cboActvtyClass' size="1">
<%
        sqlStr = "SELECT * FROM [ActvtyClassList]"
        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rsEdit.EOF Or rsEdit.BOF) Then
                rsEdit.MoveFirst
                Do Until rsEdit.EOF
                		if rsEdit.Fields("ActvtyClass_DPA_") = rs.Fields("ActvtyClass_DPA_") Then%>
                			<option selected value = '<%=rsEdit.Fields("ActvtyClass_DPA_")%>'><%=rsEdit.Fields("ActvtyClassDescription")%></option>
                		<%else%>
                        <option value = '<%=rsEdit.Fields("ActvtyClass_DPA_")%>'><%=rsEdit.Fields("ActvtyClassDescription")%></option>
                     <%end if
						rsEdit.MoveNext
                Loop
        End If
    
%>

    </select></td>
  </tr>
  <tr>
  <td width="17%"></td>
  <td width="83%"><select name="cboStatus">	
			<% if rs("status")=1 then%>		
			<option selected SearchCode = "0" SearchText = "Pending" value = '1'>Pending</option>			
			<option value='2'>In progress</option>
			<option value='3'>Resolved</option>
			<% end if %>
			<% if rs("status")=2 then%>		
			<option SearchCode = "0" SearchText = "Pending" value = '1'>Pending</option>			
			<option selected value='2'>In progress</option>
			<option value='3'>Resolved</option>
			<% end if %>
			<% if rs("status")=3 then%>		
			<option SearchCode = "0" SearchText = "Pending" value = '1'>Pending</option>			
			<option value='2'>In progress</option>
			<option selected value='3'>Resolved</option>
			<% end if %>						
		</select>
      </td>
      
  </tr>
   <tr>
    <td width="17%">Date</td>
    <td width="83%"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
  </tr>
  
  <tr>
    <td width="17%">Notes</td>
    <td width="83%"><textarea name ='txtNotes' id = 'txtNotes' rows="5" cols="34"><%=rs.Fields("ActivityNotes")%></textarea></td>
  </tr>
  <tr>
    <td width="100%" COLSPAN=2 align="right" valign=absBottom>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">		
		&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
    </td>
  </tr>
</table>
</form>

</body>

</html>
