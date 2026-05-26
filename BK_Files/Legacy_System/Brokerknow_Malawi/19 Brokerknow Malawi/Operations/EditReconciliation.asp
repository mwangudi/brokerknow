<!--#include file="../libroutines.asp"-->
<%
	
	Dim action
	Dim conn 
   	Dim sqlStr
   	Dim rs
   	Dim ID
   	Dim ReconcileDate		 	
	action = ucase(Request.Form("action"))
	
	UserId=Session("UserID")

	ID = Request("ID")
	itemids=Split(ID,"<->")
	itemid=itemids(1)
	itemtypeid=itemids(0)
	PDate=itemids(2)
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
        
	  LevyRate=Request.Form("txtLevy")

        If cmdCancel <> "" Then
			
			WriteDialogCancelScript
			Set Conn = Nothing
			Response.End
        End If
      
        ReconcileDate=Request.Form("txtDate")

        'validate Description
        'validate Slip
				If Trim(ReconcileDate) = "" Then%>
				         <script language = 'vbscript'>
				         		ShowMessage "Please Select Reconciliation date.<%=itemid%>"				         		
				         </script>
				         <% response.end
				 End If	
        
	  		conn.BeginTrans
	  					if(Cint(itemtypeid)=3)	then
	  					sqlStr="update Payment Set ReconcileDate=#" & FormatDate(CDate(ReconcileDate)) & "# ,TimeChanged=GetDate(),ChangedBy= " & UserId & " where BankAccount_DPA_=4 and Cast(Floor(Cast(PaymentPDate as Float)) AS DateTime)=#" & PDate & "#"
	  					Else
							if(Cint(itemtypeid)=1) then				
							sqlStr="update Payment Set ReconcileDate=#" & FormatDate(CDate(ReconcileDate)) & "# ,TimeChanged=GetDate(),ChangedBy= " & UserId & " where Payment_DPA_=" & itemid
        						else
        						
							sqlStr="update JournalEntry Set ReconcileDate=#" & FormatDate(CDate(ReconcileDate)) & "# ,TimeChanged=GetDate(),ChangedBy= " & UserId & " where JournalEntry_DPA_=" & itemid
							end if
						end if
												
				sqlStr = SQLServerFormat(HandleQuote(sqlStr))	
					
                conn.Execute(sqlStr)					
	  
        conn.CommitTrans
               
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
   	end If
   	

        Set conn = GetActiveConnection("KBroker")
       
        
        sqlStr = "SELECT * FROM [ReconciliationList] WHERE Entryid  = " & itemid & " and type=" & itemtypeid
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Account does not exist"
                		window.self.close
                </script>
                <% response.end
        End If
        
        
if(Cint(itemtypeid)=1) then
AccountType="Payment"
else
AccountType="Journal"
end if
%>   
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%></title>

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
<% if(IsNull(rs("ReconcileDate")) or rs("ReconcileDate")="") then 
	ReconcileDate=Date
	else
	ReconcileDate=CDate(rs("ReconcileDate"))
	end if
%>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmEditActvtyClass", "txtDate","cmdDate","<%=  FormatDate(ReconcileDate) %>",1);
</SCRIPT>

<form name = 'frmEditActvtyClass' id="frmMain" method = 'post' action = 'EditReconciliation.asp' >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="40%">Account Name</td>
    <td width="60%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtAccount' id = 'txtAccount' size="20" readonly value = '<%=rs.Fields("AccountName")%>'></td>
  </tr>
  <tr>
    <td width="40%">Account Type</td>
    <td width="60%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtLevy' id = 'txtLevy' size="20" value = '<%=AccountType%>'></td>
  </tr>
  <tr>
    <td width="40%">Transaction Date</td>
    <td width="60%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtAccount' id = 'txtAccount' size="20" readonly value = '<%=FormatDate(rs.Fields("Date"))%>'></td>
  </tr>
  <tr>
    <td width="40%">Reference</td>
    <td width="60%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtLevy' id = 'txtLevy' size="20" value = '<%=rs("Reference")%>'></td>
  </tr>
  <tr>
    <td width="40%">Particulars</td>
    <td width="60%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtAccount' id = 'txtAccount' size="20" readonly value = '<%=rs.Fields("Particulars")%>'></td>
  </tr>
  <tr>
    <td width="40%">Debit</td>
    <td width="60%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtLevy' id = 'txtLevy' size="20" value = '<%=FormatNum(rs("Debit"))%>'></td>
  </tr>
  <tr>
    <td width="40%">Credit</td>
    <td width="60%">&nbsp;&nbsp;<input readonly = 'true' class=readonly  STYLE="WIDTH: 200px; text-align: left" type = 'text'  name ='txtAccount' id = 'txtAccount' size="20" readonly value = '<%=FormatNum(rs.Fields("Credit"))%>'></td>
  </tr>
  <tr>
    <td width="40%">Reconciliation Date</td>
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
