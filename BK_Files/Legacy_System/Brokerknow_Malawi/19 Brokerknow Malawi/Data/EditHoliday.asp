<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Edit Holiday</title>

  <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
 <script language="JavaScript" src="CALENDAR/calendar.js"></script>
  <!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">

<script language='javascript'>
		function forceSubmit()
		{
			setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frmEditHoliday.method='post';
			document.frmEditHoliday.target='_self';
			document.frmEditHoliday.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}

</script>
</head>

<body Class="Dialog" onload="setOpener()">
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
		
        
       name = Request.Form("txtName") 
       HolidayDate = Request.Form("txtHolidayDate")
             
       'validate
        If Trim(name) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the description"
                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
   
        'validate 
        If Trim(HolidayDate) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the holiday date"
                 
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
  
        
        Set conn = GetActiveConnection("KBroker")
        
        'save data
       sqlStr = "UPDATE [Holidays] SET Holiday = " & "#" & FormatDate(HolidayDate) & "#" & ", Description = " & "'" & name & "'" & "" & _
				" WHERE Holiday_DPA_  = " & ID
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript2
        Response.End
   	end If
   	
   	 Set conn = GetActiveConnection("KBroker")
   	
   	 sqlStr = "SELECT * FROM [HolidayList] WHERE Holiday_DPA_ = " & ID
     Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
     
     If rsEdit.EOF Or rsEdit.BOF Then    
     %>	
		<Script Language="JavaScript">
			ShowMessage("The holiday cannot be retrieved for editing.")
		</Script>
      
      <%Set rsEdit = Nothing
		Set Conn = Nothing				
		WriteDialogRefuseOpenScript
		Response.End
     End If   
%>

<form name = 'frmEditHoliday' method = 'post' action = 'EditHoliday.asp' >
<table border="0" width="100%">
  <tr>
    <td width="30%">Description</td>
    <td width="70%"><input type = 'text' name ='txtName' id = 'txtName' size="20" value="<%= rsEdit.Fields("Description").Value %>"></td>
  </tr>
  <tr>
	<td width="30%"><b><font color="#000080">Holiday</font></b></td>
    <td width="70%"> 
		<SCRIPT language="JavaScript">			
			var cal=new ctlSpiffyCalendarBox("cal", "frmEditHoliday", "txtHolidayDate","cmdDate","<%= FormatDate(rsEdit.Fields("Holiday").Value) %>",1);
			cal.writeControl();
		</SCRIPT>
    </td>
  </tr>
 </table>


<table border=0 width="100%">  
  <tr>
     <td align=right>
		<BR>
		<BR>
		<BR>
		<BR>
		<input type = 'button' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value="Save" onclick="forceSubmit();">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">		
		<input type = 'hidden' name ='action' id = 'action' value="Execute">&nbsp;
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
     </td>
  </tr>
</table>
</form>
</body>

</html>














