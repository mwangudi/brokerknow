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
useOwner = Request.Form("useOwner")
selOwners = Request.Form("selOwners")

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
	<form method="POST" action="TradeAffirmationBond.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<input type="hidden" Name="timeLimit" value="1">
		<table>
			
			<tr>
				<td colspan="2"> <input type="checkbox" OnClick="JavaScript: switchDisplay (document.all.item('ownerSelectRow')); " class="BorderLess" name="useOwner" id="useOwner" value="1"> &nbsp; &nbsp; <label for="useOwner" style="cursor: hand">Narrow down to specific owner/s (optional)</label></td>
			</tr>
			
			<tr style="display: none;" id="ownerSelectRow" align="right">
				<td colspan=2>
					<select name="AllSelOwners" style="display: none">
						<%
						Set conn = GetActiveConnection("KBroker")
						sqlStr = "SELECT DISTINCT LotTDate, Owner FROM TradeAffirmationBond"   
						Set Rs = Conn.Execute (SQLServerFormat(HandleQuote(sqlStr)))
						If Not (Rs.EOF Or Rs.BOF) Then
							Do Until Rs.EOF%>
								<option TAG="<%= FormatDate(Rs.Fields("LotTDate").value) %>" value="<%= Rs.Fields("Owner").Value %>"><%= Rs.Fields("Owner").Value %></option>
							<%Rs.MoveNext
							Loop
						End If%>								
					</select>
					
					<select name="selOwners" size="10" multiple>
											
					</select>
					<Script Language="JavaScript">
						calendarChange ("<%= FormatDate(Date) %>")						
					</Script>
				</td>
			</tr>
			
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
   
   Set TimeLimitRs = CreateObject("ADODB.Recordset")						        
   TimeLimitRs.CursorLocation = adUseClient	

   Set conn = GetActiveConnection("KBroker")
    	
   sqlStr="SELECT TimeLimitLimDaysNSE From TimeLimit where TimeLimit_DPA_=6"
   set TimeLimitRs=Conn.Execute(sqlStr) 
   
	if not(TimeLimitRS.eof and TimeLimitRs.bof) then
	 NoOfDays=TimeLimitRS("TimeLimitLimDaysNSE")
	end if
		
	SettlementDate=LTdate(CDate(selectedTradeDate),NoOfDays)	

   sqlStr = "SELECT * FROM TradeAffirmationBond WHERE LotTDate = '" & FormatDate(selectedTradeDate) & "'"
   
   If useOwner = "1" Then
		
		If selOwners <> "" Then
			selOwners = Replace(selOwners, "'", "''")
			selOwners = Replace(selOwners, ",", "','")
			selOwners = "'" & selOwners & "'"		
			sqlStr = sqlStr & " AND Owner IN (" & selOwners & ")"
		End If	
   End If
   
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
   
   sqlStr = "SELECT dbo.Db_TradeAffirmationBond.*,Client.ClientAddr FROM dbo.Db_TradeAffirmationBond" & _
            " inner join Client on dbo.Db_TradeAffirmationBond.client_DPA_=Client.Client_DPA_"
    
      
   If useOwner = "1" Then
		If selOwners <> "" Then
			sqlStr = sqlStr & " AND Owner IN (" & selOwners & ")"
		End If	
   End If   
   
   
   ''Response.write(sqlStr)
   ''Response.end
   
   Set sumRs = Conn.Execute(sqlStr)            
   
   Dim pageNumber
	
	pageNumber = 0
	Dim balance
	Dim acntManag
	Dim commission
	Dim NetProceeds
	
   Do Until sumRs.EOF  		
		 commission = sumRs.Fields("TotalCommission").Value	 		 
		
		OwnerName = sumRs.Fields("Owner").Value
		acntManag = sumRs.Fields("AccountManager").Value
		if isnull(sumRs("orderRef")) or trim(sumRs("orderRef")) = "" then
				orderRef = sumRs.Fields("Order_DPA_").Value				
		else
				orderRef = sumRs.Fields("Order_DPA_").Value & "/" & orderRef
		end if
		
		totalQty = sumRs.Fields("Quantity").Value
		pageNumber = pageNumber + 1
				
%> 

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%">
	<tr class="pageNumbering">
		<td align="left" >
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
	
	<tr>
		<td align="right">           			
        <table border="0" width="100%">
          <tr>            
            <td width="50%" align="right" height="250px">
            <!--<Img Src="../data/photos/image002.gif">-->
			<!--#Include file="Final_Logo.htm"--></td>
          </tr>
        </table>           			
		</td>		
	</tr>
	<tr>		
        <td width="50%"><B><%=FormatDate(Date)%></B></td>					
	</tr>
</table>	  

  <table border="0" cellspacing="0" cellpadding="5" style="font-family: Arial" width="100%">
	<tr>
      <td colspan="2" align="left"> <%= sumRs.Fields("Account").Value %> <BR> <%= sumRs.Fields("ClientAddr").Value %></td>
    </tr>
    <tr>
      <td colspan="2" align="left" height="16" valign="bottom">Attention:&nbsp;<%= OwnerName %> </td>
    </tr>
    <tr>
      <td colspan="2" align="center"><b><font size="4"><u> BOND TRADE
        CONFIRMATION</u></font></b></td>
    </tr>
    <tr>
		<td  colspan="2" align="right" height="16">
            &nbsp;			
		</td>		
	</tr>
    <tr>
      <td colspan="2" align="left">Further to your instructions we confirm <%= sumRs.Fields("TradeAction") %> on your behalf of the following bond</td>
    </tr>
    <tr>
		<td  colspan="2" align="right" height="16">
            &nbsp;			
		</td>		
	</tr>
    <tr>
      <td width="20%">Issue No:</td>
      <td><b><%= sumRs.Fields("BondIssue").Value %></b></td>
    </tr>
    <tr>
      <td width="20%">Deal Price:</td>
      <td><b><%= FormatNumEx(sumRs.Fields("LotPrice").Value,4) %>&nbsp;%</b></td>
    </tr>
    <tr>
      <td>Face Value:</td>
      <td><b><%= FormatNum(totalQty) %></b></td>
    </tr>
    <tr>
      <td>Consideration:</td>
      <td><b><%= FormatNum(totalQty*sumRs.Fields("LotPrice").Value/100) %></b></td>
    </tr>
    <tr>
      <td>Commission:</td>
      <td><b><%= FormatNumEx(commission,2) %></b></td>
    </tr>
    <tr>
      <td>Net Proceeds:</td>
      <td><b><%if sumRs.Fields("OrderTypeSale") = false Then%>
				<%= FormatNumEx((totalQty*sumRs.Fields("LotPrice").Value/100)+commission,2) %>
		<%else%>
				<%= FormatNumEx((totalQty*sumRs.Fields("LotPrice").Value/100)-commission,2) %>
		<%end if%>
</b></td>
    </tr>
    <tr>
      <td width="20%">Issue date:</td>
      <td><b><%= FormatDate(sumRs.Fields("IssueDate").Value) %></b></td>
    </tr>        
    <tr>
      <td width="20%">Settlement date:</td>
      <td><b>&nbsp;</b></td>
    </tr>  
    <%'' FormatDate(SettlementDate) %>
    <tr>
      <td width="20%">Maturity date:</td>
      <td><b><%= FormatDate(sumRs.Fields("MaturityDate").Value) %></b></td>
    </tr>
    <tr>
      <td>Counter Party:</td>
      <td><b><%= sumRs.Fields("BrokerName").Value %></b></td>
    </tr>
        
  </table>
<br>  
<table border="0" cellspacing=2 cellpadding=2 width="100%">
	<tr>
		<td align="left" valign="top">
<PRE><font face="Arial"><b>
Yours faithfully, <BR>	
FOR: DYER AND BLAIR LIMITED</b>
</font></PRE>			
		</td>		
	</tr>
	<tr>
		<td align="left" valign="bottom" height="70">
<PRE><font face="Arial"><b>		
<%= acntManag %> <BR> RELATIONSHIP MANAGER</b>
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