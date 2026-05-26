<!--#include file="../libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
   	Dim sqlStr
   	Dim rs
   	Dim ID
   	Dim ReconcileDate		 	
	action = ucase(Request.Form("action"))
	
	
	ID = Request("ID")
	itemids=Split(ID,"<->")
	itemid=itemids(1)
	TableName=itemids(0)

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		window.self.close
                </script>
                <% response.end
        End If

	if action = "EXECUTE" then
	  
	  cmdCancel = Request.Form("cmdCancel")
        Set conn = GetActiveConnection("KBroker")  
	  

        If cmdCancel <> "" Then
			
			WriteDialogCancelScript
			Set Conn = Nothing
			Response.End
        End If
      
        RegistrationDate=Request.Form("txtDate")
		OpeningBalance=Request.Form("txtOpeningBalance")
        
        'validate Registration        
		If Trim(RegistrationDate) = "" Then%>
		         <script language = 'vbscript'>
		         		ShowMessage "Please Select Registration Date"				         		
		         </script>
		         <% response.end
		 End If	
        
        'validate Opening Balance
		If Trim(OpeningBalance) = "" Then%>
		    <script language = 'vbscript'>
		         	ShowMessage "Please specify the Opening Balance "					         	
		    </script>
		    <% response.end
		End If		
		
		OpeningBalance=Replace(OpeningBalance,",","")
		
		'ensure Amount is numeric
		If (OpeningBalance <> "") And (Not IsNumeric(OpeningBalance)) Then%>
		    <script language = 'vbscript'>
				ShowMessage "Opening Balance must be numeric"
							
		    </script>
		    <% response.end
		End If	
        
		OpeningBalance=Ccur(OpeningBalance)
				
	  		conn.BeginTrans				
			
			if(TableName="Account") then
			sqlStr="update " & TableName & " Set " & trim(TableName) & "OpeningDate ='" & FormatDate(CDate(RegistrationDate)) & "'," & TableName & "OpeningBal =" & OpeningBalance & " where " & TableName & "_DPA_=" & itemid										
			else
			sqlStr="update " & TableName & " Set " & trim(TableName) & "RegDate ='" & FormatDate(CDate(RegistrationDate)) & "'," & TableName & "OpeningBal =" & OpeningBalance & " where " & TableName & "_DPA_=" & itemid							
			end if

            conn.Execute(sqlStr)								
        conn.CommitTrans
               
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
   	end If
   	

        Set conn = GetActiveConnection("KBroker")
       
        
        sqlStr = "SELECT * FROM [OpeningBalances] WHERE Entity_DPA_  = " & itemid & " and TableName=('" & trim(TableName) & "')"
        'Response.write(sqlStr)
        'Response.End
        
        
        Set rs = conn.Execute(sqlStr)
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Account does not exist"
                		window.self.close
                </script>
                <% response.end
        End If      
        
%>   
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Opening Balances</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
 <SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

</head>
<BR>
<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmEditActvtyClass", "txtDate","cmdDate","<%=  FormatDate(rs.Fields("RegistrationDate")) %>",1);
</SCRIPT>

<form name = 'frmEditActvtyClass' id="frmMain" method = 'post' action = 'EditOpeningBalance.asp' >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="40%">Entity</td>
    <td width="60%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtEntity' id = 'txtEntity' size="20" readonly value = '<%=rs.Fields("EntityType")%>'></td>
  </tr>
  <tr>
    <td width="40%">Account</td>
    <td width="60%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtAccount' id = 'txtAccount' size="20" value = '<%=rs("EntityName")%>'></td>
  </tr>
  <tr>
    <td width="40%">Code</td>
    <td width="60%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtCode' id = 'txtCode' size="20" value = '<%=rs("EntityCode")%>'></td>
  </tr>
  <tr>
    <td width="40%">Opening Balance</td>
    <td width="60%">&nbsp;&nbsp;<input  type = 'text'  name ='txtOpeningBalance' id = 'txtOpeningBalance' size="20" value = '<%=FormatNum(rs.Fields("OpeningBalance"))%>' onchange='format2Number(this);'></td>
  </tr>
  <tr>
    <td width="40%">Registration Date</td>
    <td width="60%">&nbsp;&nbsp;<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
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
    </td>
  </tr>
</table>
</form>

</body>

</html>
