<html>

<head>
	<meta http-equiv="Content-Language" content="en-uk">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	
	<title>Trade Confirmation</title>
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
			
			
			margin-left: 5cm;
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
	<form method="POST" action="DB_TradeAffirmation.asp" Name="frmMain" id="frmMain">
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
						sqlStr = "SELECT DISTINCT LotTDate, Owner FROM DB_Affirmation"   
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
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp; </td>
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
   Dim ListRs
   
   Set TimeLimitRs = CreateObject("ADODB.Recordset")   						        
   TimeLimitRs.CursorLocation = adUseClient	
	
   Set ListRs = CreateObject("ADODB.Recordset")   						        
   ListRs.CursorLocation = adUseClient	
	

Set conn = GetActiveConnection("KBroker") 
   	   
   sqlStr = "SELECT * FROM DB_Affirmation WHERE LotTDate = '" & FormatDate(selectedTradeDate) & "'"
   
   'Response.write(sqlStr)
   'Response.end
   
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
      
   sqlStr = "SELECT distinct Order_DPA_,BalanceQty,OrderRef,AccountManager,Class_DPA_ FROM DB_Affirmation WHERE LotTDate = '" & FormatDate(selectedTradeDate) & "'"
   
   'Response.write(sqlStr)
   'Response.end
    
   If useOwner = "1" Then
		If selOwners <> "" Then
			sqlStr = sqlStr & " AND Owner IN (" & selOwners & ")"
		End If	
   End If
   
   'sqlStr = sqlStr & "	GROUP BY OrderRef, Order_DPA_,BalanceQty"
   
   'Response.Write(sqlStr)
   'Response.End 
   
   Set sumRs = Conn.Execute(sqlStr)         
   Dim pageNumber
	
	pageNumber = 0
	Dim balance
	Dim acntManag
	Dim Price
	Dim TotalGross  
   
   Do Until sumRs.EOF
   		totalQty=0
   		settlementAmt=0
   		TotalGross=0
   		
   		if(sumRs("Class_DPA_") =6)	 then
         sqlStr="SELECT TimeLimitLimDaysNSE From TimeLimit where TimeLimit_DPA_=5"   
        else
   		sqlStr="SELECT TimeLimitLimDaysNSE From TimeLimit where TimeLimit_DPA_=1"   
   		end if
   
   		set TimeLimitRs=Conn.Execute(sqlStr) 
   
		if not(TimeLimitRS.eof and TimeLimitRs.bof) then
	 	NoOfDays=TimeLimitRS("TimeLimitLimDaysNSE")
		end if
	
		SettleDate=LTdate(CDate(selectedTradeDate),NoOfDays)	    

	'Response.write(SettleDate)
   	'Response.end   

        sqlStr1="SELECT distinct SettlementAmount,LotQty, LotSlipNo, SettlementGrossAmount, LotPrice FROM DB_Affirmation WHERE( Order_DPA_ = " & sumRs("Order_DPA_") & ") and LotTDate = '" & FormatDate(selectedTradeDate) & "'"
		
		'Response.write(sqlStr)
		'Response.end 
		
		Set TotalRs = CreateObject("ADODB.Recordset")
        TotalRs.CursorLocation = adUseClient
        Set TotalRs = Conn.Execute(sqlStr1)
       
        do while TotalRs.eof=false
        
        totalQty = totalQty+TotalRs.Fields("LotQty").Value
		settlementAmt = settlementAmt + TotalRs.Fields("SettlementAmount").Value	
		TotalGross=TotalGross+TotalRs.Fields("SettlementGrossAmount").Value	

		
		TotalRs.moveNext
        loop

       set TotalRs=nothing 		
       
		OrderRef = sumRs.Fields("OrderRef").Value	
		 balance = sumRs.Fields("BalanceQty").Value
		 
		sqlStr = "SELECT DB_Affirmation.*,'' as AccountManager FROM DB_Affirmation WHERE LotTDate = '" & FormatDate(selectedTradeDate) & "' And OrderRef = '" & OrderRef & "' AND Order_DPA_ = " & sumRs.Fields("Order_DPA_").Value		
		Set Rs = Conn.Execute (SQLServerFormat(HandleQuote(sqlStr)))
		OwnerName = Rs.Fields("Owner").Value
		acntManag = Rs.Fields("AccountManager").Value
		if isnull(orderRef) or trim(orderRef) = "" then
				orderRef = sumRs.Fields("Order_DPA_").Value
		else
				orderRef = sumRs.Fields("Order_DPA_").Value & "/" & orderRef
		end if
		
		pageNumber = pageNumber + 1
		
		if(Trim(rs("ClassDescription"))="Other Fund Managers") then
		SettleDate=LTdate(CDate(selectedTradeDate),3)
		end if
		
		if(isnull(Sumrs("AccountManager")) or Sumrs("AccountManager")="") then
		AccountManager="Paul Nyaga"
		else
		AccountManager=Sumrs("AccountManager")
		end if
		'AccountManager="Paul Nyaga"
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
            <td align="right" >
            <!--<Img Src="../data/photos/image002.gif">-->
			<Img Src="../data/photos/aaprintlogo.bmp">			
          </tr>
        </table>
			
						
           			
		</td>		
	</tr>
	<tr>
		<td width="50%" valign="bottom" align="left"><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=FormatDate(Rs("LotTDate"))%></B></td>
	</tr>
</table>	  

  <table border="0" cellspacing="4" cellpadding="5" style="font-family: Arial" width="100%">

    <tr>
      <td colspan="2" align="center"><b><font size="4">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<u>TRADE
        CONFIRMATION</u></font></b></td>
    </tr>
    <tr>
      <td width="20%" align="left"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;TO:</b></td>
      <td><b><%= OwnerName %></b></td>
    </tr>
    <tr>
      <td width="20%" align="left"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;FAX NO:</b></td>
      <td><b><%= Rs.Fields("ClientFax").Value %></b></td>
    </tr>
    <tr>
      <td width="20%" align="left"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;DATE:</b></td>
      <td><b><%=FormatDate(Rs("LotTDate")) %></b></td>
    </tr>
    <tr>
      <td colspan="2" align="left">&nbsp;</td>
    </tr>
    <tr>
      <td width="20%" align="left"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<u>RE:</u><b></td>
      <td><b><u>
		<%= Rs.Fields("ReferenceHeader")  %>
      </u></b></td>
    </tr>
    <tr>
      <td colspan="2" align="left">&nbsp;</td>
    </tr>
    <tr>
      <td><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;TRADE DATE:</b></td>
      <td><b><%= FormatDate(selectedTradeDate) %></b></td>
    </tr>
    <tr>
      <td><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;SECURITY TRADED:</b></td>
      <td><b><%= Rs.Fields("OrdDetailSecurity").Value %></b></td>
    </tr>
    <tr>
	<% if(rs("OrderTypeSale")=true) then
		ShareTerm="SOLD"
		else
            ShareTerm="BOUGHT"
	   end if	
	%>
      <td><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;SHARES&nbsp;<%=ShareTerm%>:</b></td>
      <td><b><%= FormatNumCommasOnly(totalQty) %></b></td>
    </tr>
    <tr>
      <td width="20%"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;DEAL PRICE:</b></td>
      <td><b>SEE OVERLEAF</b></td>
    </tr>
    <tr>
      <td width="40%"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;SETTLEMENT AMOUNT:</b></td>
      <td><b><%= FormatNum(RoundPoint05(settlementAmt)) %></b></td>
    </tr>
    <tr>
      <td><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;SETTLEMENT DATE:</b></td>
      <td><b><%= FormatDate(SettleDate) %></b></td>
    </tr>    
    <tr>
      <td><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CONTRACT NO:</b></td>
      <td><b><%= Rs.Fields("ContractNumber").Value %></b></td>
    </tr>
    <tr>
      <td width="20%"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CUSTODIAN:</b></td>          
	  <td><b>[<%= Rs("Client_DPA_") %>]&nbsp;<%= Rs.Fields("Account").Value %> <br>								
	  <%= Rs("AccountAddress") %> <br>			
	  </b></td>
	</tr>	  
	<tr>
      <td><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CDS ACCOUNT NO:</b></td>
      <td><b><%= Rs("ClientCDSNo") %></b></td>
    </tr>
    <tr>
      <td><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;OUTSTANDING BALANCE:</b></td>
      <td><b><%= iif (balance = 0,"Nil",FormatNumCommasOnly(balance)) %></b></td>
    </tr>
    <tr>
      <td><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ORDER NO:<b></td>
      <td><b><%= rs("OrderRef") %></b></td>
    </tr>   
  </table>
<br>  
<table border="0" cellspacing=2 cellpadding=2 width="100%">
	<tr>		
		<td align="left" valign="top" width="80%">
<PRE><font face="Arial"><b>		
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;for Dyer & Blair Investment Bank Ltd.</b>
</font></PRE>					
	</tr>
	<tr>
		<td  align="right" height="30">
            &nbsp;			
		</td>		
	</tr>
	<tr>
		<td align="left" valign="top">
<PRE><font face="Arial"><b>		
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=AccountManager%></b>
</font></PRE>			
		</td>		
	</tr>
	<tr>
		<td align="left" valign="bottom" height="70">
<PRE><font face="Arial"><b>		
<%= acntManag %></b>
</font></PRE>			
		</td>		
	</tr>	
	<tr>
		<td width="100%"><!--#Include file="DirectorFooter.asp"-->&nbsp;</td>
	</tr>
</table>  


<BR class="newpage">
<%
pageNumber = pageNumber + 1
%>
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%">
	<tr class="pageNumbering">
		<td align="left" >
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>

	<tr>
		<td align="right"  style="border-bottom: 2px inset #000000">
			
						
           			
        <table border="0" width="100%">
          <tr>
            <td width="50%" valign="bottom"><B><%=FormatDate(selectedTradeDate)%></B></td>
            <td width="50%" align="right" height="250px"><Img Src="../data/photos/aaprintlogo.jpg"></td>
          </tr>
        </table>
			
						
           			
		</td>		
	</tr>
	<tr>
		<td align="right" height="16">
            &nbsp;			
		</td>		
	</tr>
</table>	  
<table border="0" cellspacing="0" cellpadding="5" style="font-family: Arial" width="100%">

    <tr>
    	<td>&nbsp;</td>
		<td>&nbsp;</td>
      	<td colspan="5" align="center"><b><font size="4"><u>TRANSACTION LIST</u></font></b></td>
    </tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
      <td colspan="5" align="left" width="80%"><b><font size="3">The Quantities traded are broken down as follows: </font></b></td>
    </tr>
    	
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td><u><b>SLIP NO</b></u></td>
    	<td><u><b>Quantity</b></u></td>
    	<td><u><b>Deal Price</b></u></td>
    	<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<%
	sqlStr="Select LotQty,LotSlipNo,LotPrice From Lotlist Where(Client_DPA_=" & Rs("Client_DPA_") & ") and (Security_DPA_=" & Rs("Security_DPA_") & ")" & _
			"and (LotTDate='" & Rs("LotTDate") & "') and (Order_DPA_=" & Rs("Order_DPA_")& ")"
			
	'Response.write(sqlStr)
	'Response.end
	
	Set ListRs=Conn.Execute(sqlStr)
	
	Do while ListRs.eof=false
	%>
	<tr>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td><%=ListRs("LotSlipNo")%></td>
    <td><%=FormatNumEx(ListRs("LotQty"),0)%></td>
    <td><%=FormatNum(ListRs("LotPrice"))%></td>
    <td>&nbsp;</td>
	<td>&nbsp;</td>
    </tr>
	<%
	ListRs.MoveNext
	loop
	
	%>
</table>

<%		sumRs.MoveNext
		'important!		
			If Not sumRs.EOF Then %>
				<BR class="newpage">				 
		<%	
		'pageNumber = pageNumber + 1
		End If
	Loop

	Set sumRs = Nothing
	Set Rs = Nothing
Set Conn = Nothing%>
</body>

</html>