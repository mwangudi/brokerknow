<html>

<head>
	<meta http-equiv="Content-Language" content="en-uk">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	
	<title>Bond Trade Confirmation</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	  <!--CALENDAR -->
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
    <style media="print">
		@page {
			@top{font-family: Helvetica, Arial, sans-serif;
				font-size: 150%;
				font-weight: bolder;
				text-align: left;
				content: "<%= FormatDate(Date) %>";			
			}		
			
			
			margin-left: 2cm;
			margin-right: 5cm;
			margin-top: 1cm;    
			margin-bottom: 2cm;
			size: portrait;
			
			br.newpage{
				page-break-before:always;
			}
			
			tr.pageNumbering{
				display:none;
			}
			
		}

	</style>
</head>

<body Class="Reports">

<!--#include file="../libroutines.asp"-->


<%

genReport = Request.Form("genReport")
selectedTradeDate = Request.Form("txtDate")
'useOwner = Request.Form("useOwner")
'selOwners = Request.Form("selOwners")

If genReport <> "1" Or Not IsDate(selectedTradeDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
	
		function validateForm(frm){			
			if (frm.txtDate.value==''){
				alert("Select a date");
				frm.txtDate.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		function switchDisplay(obj){
			if (obj.style.display=='none') obj.style.display = '';
			else obj.style.display = 'none';
		}
		
		
		function calendarChange(dateValue){
			var Input = document.all.item('AllSelOwners');
			var Output =  document.all.item('selOwners');
			var dateVal;
			var mainObj = document.all.item('txtDate');
			
			if (dateValue==null || dateValue=='undefined') dateVal = mainObj.value;
			else dateVal = dateValue;
			
			Output.length = 0;
		     for (loop=0; loop < Input.length; loop++){
		     		if (Input.options[loop].TAG == clientFormatDate(dateVal)){		     			
		     		    NewOption = new Option();   			    
		   			    NewOption.text = Input.options[loop].text;
		   			    NewOption.value = Input.options[loop].value;			
		   			    Output.add(NewOption, 0);
		     		}	
		     }
		}
		
		
		var cal = new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
		
		
	</Script>
	
	<Script Language="VBScript">
		Function clientFormatDate(theDate)
			On Error Resume Next
			clientFormatDate = Day(theDate) & "-" & MonthName(Month(theDate), True) & "-" & Year(theDate)
			If Err.Number > 0 Then
				clientFormatDate = theDate
			End If
		End Function
	</Script>
	
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="BondTradeConfirmation.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<input type="hidden" Name="timeLimit" value="1">
		<table>		
			
			<tr>
				<td>Select any day of month</td>
				<td>
					<SCRIPT language="JavaScript">
						cal.writeControl();
					</SCRIPT>	
				</td>
			</tr>
			
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp; </td>
			</tr>
		</table>
		
	</form>
		
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>

<% DrawPageFunctions True, True, True %>

<%
   Dim conn 
   Dim sqlStr
   Dim rs
   Dim isDifferentBroker
   Dim currBrokerCode
   
   Set conn = GetActiveConnection("KBroker") 	

   sqlStr = "SELECT * FROM BondConfirmationList WHERE TransDate = '" & FormatDate(selectedTradeDate) & "'"   
   
   Set Rs = Conn.Execute (sqlStr)
   If Rs.EOF Or Rs.BOF Then%>
		<Script Language="JavaScript">	
			ShowMessage('There were no traded items under the selected date');
			window.parent.history.go(-1)
		</Script>
		<%Set Conn = Nothing
		Set Rs = Nothing
		Response.End
   End If   
   
   sqlStr = "SELECT Client_2.ClientAddr AS Address, Client_2.ClientPAddr AS Physical, BondConfirmationList.*, BondConfirmation.*," & _ 
            " Owner.OwnerFname + ' ' + Owner.OwnerLName AS AccountManager, Client_1.ClientName AS CounterParty," & _ 
            " Class.AgentStatus AS AgentReturnable FROM Class INNER JOIN BondConfirmationList INNER JOIN Client Client_2 ON BondConfirmationList.Client_DPA_ = Client_2.Client_DPA_ INNER JOIN" & _
            " Owner ON Client_2.Owner_DPA_ = Owner.Owner_DPA_ ON Class.Class_DPA_ = Client_2.Class_DPA_ LEFT OUTER JOIN" & _
            " Client Client_1 Right outer JOIN BondConfirmation ON Client_1.Client_DPA_ = BondConfirmation.CounterParty_DPA_ ON " & _
            " BondConfirmationList.Order_DPA_ = BondConfirmation.Order_DPA_ WHERE TransDate = '" & FormatDate(selectedTradeDate) & "'"   
  
   
   
   'Response.write(sqlStr)
   'Response.end
   
   Set sumRs = Conn.Execute(sqlStr)            
   
   Dim pageNumber
	
	pageNumber = 0
	Dim balance
	Dim acntManag
	Dim commission
	Dim NetProceeds
	
   Do Until sumRs.EOF  				
		pageNumber = pageNumber + 1
'Response.write(sumRs("SettlementDate"))
'Response.end
				
%> 
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%">
	<tr class="pageNumbering">
		<td align="left" >
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
	
	<tr>
		<td align="left">           			
        <table border="0" width="100%">
          <tr>            
            <td align="right" colspan=2 height="200px" style="border-bottom: 2px inset #000000" width="670">
			<Img Src="../data/photos/aaprintlogo.bmp">			
		</td>		
          </tr>
        </table>           			
		</td>		
	</tr>	  	
  <tr>
  <td>
  <p class=MsoNormal style='margin-left:54.0pt'>  
  <table border="0" cellspacing="0" cellpadding="3" style="font-family: Arial" width="80%">
	<tr>		
        <td width="50%"><B><%=FormatDate(Date)%></B></td><td>&nbsp;</td>					
	</tr>      
	<tr>
      <td colspan="2" align="left"><b>
	<% if(sumRs("Title")="" or isnull(sumRs("Title"))) then
	else
	%>
	<%=ucase(sumRs("Title"))%><BR>	
	<% end if%>          
	<%= sumRs.Fields("ClientName").Value %> <BR> 
	<% if(sumRs("Physical")="" or isnull(sumRs("Physical"))) then
		else
		%>
	<%= sumRs.Fields("Physical").Value %>/
	<% end if%>

	<%= sumRs.Fields("Address").Value %></b></td>
    </tr>
    <tr>
      <td colspan="2" align="left" height="16" valign="bottom"><b>ATTENTION:&nbsp;<%= ucase(sumRs("ClientContact")) %> </b></td>
    </tr>
    <tr>
      <td colspan="2" align="left"><b><font size="3"><u><%=sumRs("Type")%>&nbsp;BOND TRADE
        CONFIRMATION</u></font></b></td>
    </tr>    
    <tr>
		<td  colspan="2" align="right" height="16">
            &nbsp;			
		</td>		
	</tr>
    <tr>
      <td width="50%">Issue No:</td>
      <td align="right"><b><%= sumRs.Fields("BondIssue").Value %></b></td>
    </tr>
    <tr>
      <td width="50%">Deal Price:</td>
      <td align="right"><b><%= FormatNumEx(sumRs.Fields("Price").Value,4) %>&nbsp;%</b></td>
    </tr>
    <tr>
      <td>Face Value:</td>
      <td align="right"><b><%= FormatNum(sumRs("FaceValue")) %></b></td>
    </tr>
    <% if(sumRs("Type"))="SALE" then		
	 else
	%>
	<tr><td>&nbsp;</td><tr>    
	<%
	end if
     %>    
    <tr>
      <td>Consideration</td>
      <td align="right"><b><%= FormatNum(sumRs("Gross")) %></b></td>
    </tr>    
    <% if(sumRs("Type"))="SALE" then		
	 else
		
	 	if(isnull(sumRs("CounterParty")) or sumRs("CounterParty")="") then		
		else
    		%>
		<tr>
      	<td><b>Pay&nbsp;<%=sumRs("CounterParty")%></b></td>
      	<td align="right">&nbsp;</td>
    		</tr>    		
		<%
		end if 	
	if(isnull(sumRs("ConsiderationI")) or sumRs("ConsiderationI")="") then
	else
	%>

    <tr>
      <td><b><%=sumRs("ConsiderationI")%></b></td>
      <td>&nbsp;</td>
    </tr> 
	<% end if
	end if
	%>   
    <tr><td>&nbsp;</td><tr>    
    <tr>
      <td>Commission:</td>
	<% if(sumRs("AgentReturnable")=true) then
		%>
      <td align="right"><b><%= FormatNum((sumRs("Commission")-sumRs("AgentCommission"))) %></b></td>
	<% 
		else
	%>
      <td align="right"><b><%= FormatNum(sumRs("Commission")) %></b></td>
	<% end if %>
    </tr>
    <% if(isnull(sumRs("CommissionI")) or sumRs("CommissionI")="") then
	else %>    
    <tr>
      <td><b><%=sumRs("CommissionI")%></b></td>
      <td>&nbsp;</td>
    </tr>
    <tr><td>&nbsp;</td><tr>    
    <% end if
	if(sumRs("Type"))="SALE" then
	%>
    <tr>    
      <td>Net amount due:</td>
      <td align="right"><b><%=formatnum(sumRs("Netamount"))%></b></td>
    </tr>
	<%
	else
	%>
    <tr>
      <td>Total cost of buyer:</td>
      <td align="right"><b><%=formatnum(sumRs("Netamount"))%></b></td>
    </tr>
    <% end if %>	
    <tr>
      <td width="50%">Trade date:</td>
      <td align="right"><b><%= FormatDate(sumRs.Fields("TransDate").Value) %></b></td>
    </tr>            
    <tr>
      <td width="50%">Issue date:</td>
      <td align="right"><b><%= FormatDate(sumRs.Fields("IssueDate").Value) %></b></td>
    </tr>        
    <tr>
      <td width="50%">Settlement date:</td>
      <td align="right"><b><%=FormatDate(sumRs("SettlementDate")) %></b></td>
    </tr>      
    <tr>
      <td width="50%">Maturity date:</td>
      <td align="right"><b><%= FormatDate(sumRs.Fields("MaturityDate").Value) %></b></td>
    </tr>
    <tr>
      <td>Counter Party:</td>
      <td align="right"><b><%= sumRs.Fields("CounterParty").Value %></b></td>
    </tr>
        
  </table>
<br>  
<table border="0" cellspacing=2 cellpadding=2 width="80%">
	<tr>
		<td align="left" valign="top">
<PRE><font face="Arial"><b>
Yours faithfully, <BR>	
FOR: DYER AND BLAIR INVESTMENT BANK LIMITED</b>
</font></PRE>			
		</td>		
	</tr>
	<tr>
		<td align="left" valign="bottom" height="70">
<PRE><font face="Arial"><b>
<%if(sumRs("AccountManager")="" or isnull(sumRs("AccountManager"))) then
%>
MARTIN MBUGUA
<% else%>		
<%= ucase(sumRs("AccountManager"))%>
<% end if%>
RELATIONSHIP MANAGER</b>
</font></PRE>			
		</td>		
	</tr>
	
	<tr>
		<td  align="right" height="30">
            &nbsp;			
		</td>		
	</tr>
	<tr>
		<td width="100%"><!--#Include file="DirectorFooter.asp"--></td>
	</tr>
</table>
</td> 
</tr> 
</table>	


<%		sumRs.MoveNext
		'important!		
			If Not sumRs.EOF Then %>
				<BR class="newpage">
		<%	End If
	Loop

	Set sumRs = Nothing
	Set Rs = Nothing
Set Conn = Nothing%>
</body>

</html>