<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Add Holiday</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
 <script language="JavaScript" src="CALENDAR/calendar.js"></script>
<script language='javascript'>
		function forceSubmit()
		{
			setOpener();
			//var targetPage = window.dialogArguments.opener.document.all.item("frmMain").elements("EditPage").value;
					
			document.frmAddHoliday.method='post';
			document.frmAddHoliday.target='_self';
			document.frmAddHoliday.submit();		
		}
		
		function setOpener()
		{
			window.self.opener = window.dialogArguments.opener;					
		}

</script>
 
</head>

<body Class="Dialog" onload="setOpener()">


<!--#include file="../libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim guidStr 
   Dim guid 
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim name
		Dim HolidayDate
		        
       name = Request.Form("txtName")
       HolidayDate = Request.Form("txtHolidayDate")
      
        'validate Name
        If Trim(name) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Holiday name"
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If
        'validate size of Name
        If Len(Name) > 100 Then%>
                		<script language = 'vbscript'>
						ShowMessage "Name can only be 100 characters in length"
                
                </script>
                <% 
				ReloadPage(ID)
				response.end
        End If      		
		'validate detail info
                'validate date
                If Trim(HolidayDate) = "" Then%>
                		<script language = 'vbscript'>
                				ShowMessage "Please specify the holiday date"
                				
                		</script>
                		<% 
						ReloadPage(ID)
						response.end
                End If
        
                

         sqlStr = "INSERT INTO Holidays (Description, Holiday_DPA_, Holiday) SELECT " & "'" & name & "'" & " as Description," & " " & "iif(isnull(max([Holiday_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Holidays'),max([Holiday_DPA_]) + 1)" & " " & " as Holiday_DPA_" & _
                "," & "'" & FormatDate(HolidayDate) & "'" & " as Holiday FROM Holidays"
        Set conn = GetActiveConnection("KBroker")
        
        'conn.BeginTrans
			conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
		'conn.CommitTrans
		Set Conn = Nothing
		WritefraEnabledDialogCloseScript2
        Response.End        
   	end If
%>

<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<form name = 'frmAddHoliday' method = 'post' action = 'AddHoliday.asp' target="deleteFrame" OnSubmit="JavaScript: UpdateDialogHandle();">
<table border="0" width="100%">
  <tr>
    <td width="30%">Description</td>
    <td width="70%"><input type = 'text' name ='txtName' id = 'txtName' size="20"></td>
  </tr>
  <tr>
	<td width="30%"><b><font color="#000080">Holiday</font></b></td>
    <td width="70%"> 
		<SCRIPT language="JavaScript">			
			var cal=new ctlSpiffyCalendarBox("cal", "frmAddHoliday", "txtHolidayDate","cmdDate","<%= FormatDate(Date) %>",1);
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
     </td>
  </tr>
</table>
</form>
</body>

</html>
