<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Payment Voucher</title>
  
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
			
			margin-left: 0cm;
			margin-right: 0cm;
			margin-top: 2cm;    
			margin-bottom: 0cm;
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
timeLimit = Request.Form("timeLimit")

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
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="PaymentVoucherByTradeDate.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<table>
			<tr>
				<td>Select any day of month</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td>
					<p><input type="radio" class="BorderLess" checked Name="timeLimit" value="0" id="NSETime"><label for="NSETime" style="cursor: hand">Use NSE time limit</label></p>
					<p><input type="radio" class="BorderLess" Name="timeLimit" value="1" id="InternalTime"><label for="InternalTime" style="cursor: hand">Use internal time limit</label></p>
					
				</td>
			</tr>
			<tr>
				<td colspan=2>&nbsp;</td>
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
   Dim rsNames
   Dim sqlStr1
   
   Set TimeLimitRs = CreateObject("ADODB.Recordset")						        
    TimeLimitRs.CursorLocation = adUseClient	

    Set conn = GetActiveConnection("KBroker")
       
	'Update the settlement Dates 
	Conn.Execute("Delete From PaymentVoucherSettlementDates")
   		sqlStr="Select * From PaymentVoucherDates" 
   		set rs=conn.Execute(sqlStr)
   
   		if not(rs.eof or rs.bof) then
   			do while rs.eof=false
   				if(Cint(rs("Class_DPA_"))=6) then
   				sqlStr="SELECT TimeLimitLimDaysNSE From TimeLimit where TimeLimit_DPA_=7"
   			    else
   			    sqlStr="SELECT TimeLimitLimDaysNSE From TimeLimit where TimeLimit_DPA_=1"
   			    end if
   			    
   			    set TimeLimitRs=Conn.Execute(sqlStr) 
   
				if not(TimeLimitRS.eof and TimeLimitRs.bof) then
	 			NoOfDays=TimeLimitRS("TimeLimitLimDaysNSE")
				end if
		
			SettlementDate=LTdate(CDate(rs("SettlementDate")),NoOfDays)	
			
			SettlementDate=FormatDate(SettlementDate)
			
			
			Conn.Execute("Insert Into PaymentVoucherSettlementDates (Contract_DPA_,SettlementDate) Values( " & rs("Contract_DPA_") & ",'" & SettlementDate & "')")
   			rs.movenext
   			loop
   		end if     
		
		set rs=nothing
		
		sqlStr = "SELECT * FROM [SettlePaymentVoucher] WHERE Day(SettlementDate) = Day('" & FormatDate(selectedTradeDate) & "') and " & _
                     " Month(SettlementDate) = Month('" & FormatDate(selectedTradeDate) & "') and Year(SettlementDate) = Year('" & FormatDate(selectedTradeDate) & "') ORDER BY ClientName"		
		
        'Response.write(sqlStr)
        'Response.end
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.CursorLocation = adUseClient
        
        'Set rsNames = CreateObject("ADODB.Recordset")
        'rsNames.CursorLocation = adUseClient
        
        Set Rs=Conn.Execute(sqlStr)
        If rs.EOF Or rs.BOF Then
               %>
                <Script Language="JavaScript">
					alert("No payment vouchers available");
					window.parent.history.back();				
                </Script>
                <% Set Rs = Nothing
                Set Conn = Nothing
                Response.End
        End If
        
        rs.MoveFirst
        

Dim PageNumber
PageNumber=0

Do Until rs.EOF
PageNumber=PageNumber+1

	currClientCode = Rs.Fields("Client_DPA_").Value 
sqlStr1= "SELECT * FROM Client Where(Client_DPA_=" & Rs.Fields("Client_DPA_").Value & ")" 

Set rsNames=Conn.execute(sqlStr1)
%>
<table width="100%" class="ReportsTable">
	<tr class="pageNumbering">
		<td align="left" height="18">
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>

<tr><td>&nbsp;</td></tr>
</table>
<table width="89%" border="1" height="50">     
     <tr>
		     <td align=center width="100%" height="44">
                <font face="Arial Narrow" size="3">&nbsp;<span lang="en-us">PAYMENT 
                VOUCHER</span></font></td>
	</tr>
</table>
<table width=700 border="0">	
	<tr>
	<td rowspan="9" width="815">&nbsp;</td>
	<td width="435">			
	<tr><td align=left width="435"><font face="Courier" size="4">DYER AND BLAIR 
      INVESTMENT BANK LTD</font></td></tr>
	<tr><td align=left width="435"><font face="Courier" size="4">(Members of the NSE Since 1954)</font></td></tr>
	<tr><td align=left width="435"><font face="Courier" size="4">&nbsp;</font></td></tr>
	<tr><td align=left width="435"><font face="Courier" size="4">10th Floor, Loita House</b></font></td></tr>
	<tr><td align=left width="435"><font face="Courier" size="4">P.O.BOX&nbsp;45396-00100,&nbsp;Nairobi&nbsp;KENYA</b></font></td></tr>
	<tr><td align=left width="435"><font face="Courier" size="4">Tel:&nbsp;+254&nbsp;020&nbsp;3240000</b></font></td></tr>
	<tr><td align=left width="435"><font face="Courier" size="4">Fax:&nbsp;+254&nbsp;020&nbsp;218633</b></font></td></tr>
	<tr><td align=left width="435"><font face="Courier" size="4">Email:&nbsp;a<span lang="en-us">d</span>min@dyer.africaonline.co.ke</b></font></td></tr>								
	</td>
	</tr>
     
	<tr>
		     <td colspan=2 width="701">
		        <font face="Arial" size="2">&nbsp;</font></td>
	</tr>	
  
</table>

<table border="0" cellspacing="0" cellpadding="0" style="font-family: Arial Narrow" width="740">    
	<tr>
	<td colspan="2" width="740">&nbsp;</td>
	</tr>
    <tr>
      <td width="473"><font face="Courier" size="4">TO THE ACCOUNTANT</font></td>      
    </tr>
    <tr>
    <td width="473">&nbsp;</td>
    </tr>
    <tr>
    <td width="473"><font face="Courier" size="4">PLEASE PAY</font></td>    
    </tr>
    <tr>
    <td width="473">&nbsp;</td>
    </tr>
    <tr>
    <td width="473"><font face="Courier" size="3"><b><%= Rs.Fields("ClientName").Value %></b></font><font face="Courier" size="4"> &nbsp;[<%= Trim(Rs.Fields("Client_DPA_").Value)%>]</font></td>        
    </tr>
    <tr>
    <td width="473"><%= RsNames.Fields("ClientAddr").Value %></td>
    </tr>
    <tr>
    <td width="473">&nbsp;</td>
    </tr>    
    <tr>
    <td width="473"><font face="Courier" size="4">Tel:&nbsp;<%= RsNames.Fields("ClientOfficeTel").Value %></font></td>
    <td width="267"><font face="Courier" size="4">VOUCHER NO :</font></td>    
    </tr>
    <tr>
    <td width="473">&nbsp;</td>
    </tr>        
    <tr>
      <td width="473">
      <font face="Courier" size="4">&nbsp; </font>
       </td>

      <td width="267">
      <p align="left"><font size="4"  face="Courier">PAYMENT DATE:
        <%= FormatDate(rs("SettlementDate")) %></font></td>
    </tr>
    <tr>
    <td width="473">&nbsp;</td>
    </tr>
	<tr>
      <td width="473"><font face="Courier" size="4">Being Payment For :</font></td>      
    </tr>   
	<tr>
    <td width="473">&nbsp;</td>
    </tr>
	
<BR>

<!--<center>-->
<table border="0" cellspacing=0 cellpadding=0 width="650" height="266">

<tr>
	<td height="66" width="6"><font face="Courier" size="4" align="left">SLIP</font></td>
	<td height="66" width="17"><font face="Courier" size="4">CONTRACT&nbsp;</font></td>
	<td height="66" width="92"><font face="Courier" size="4">TRADED</font></td>	
	<td height="66"  width="377"><font face="Courier" size="4">SECURITY</font></td>	
	<td align=right  height="66" width="90"><font face="Courier" size="4">QUANTITY</font></td>
	<td align=right  height="66" width="200">
    <font face="Courier" size="4">&nbsp;&nbsp;PRICE</font></td>	
	<td align=right  height="66" width="150">
    <font face="Courier" size="4">&nbsp;&nbsp;NET&nbsp;AMOUNT</font></td>
	
  </tr>
<%		
	totalNetAmount = 0
	isDifferentClient = False
	
		 Do Until isDifferentClient 
			totalNetAmount = totalNetAmount + FormatNum(rs.Fields("NetAmount")) %>
                <tr>
                        <td  height="5" width="6"><%=rs.Fields("LotSlipNo")%></td>
                        <td  height="5" width="12"><%=rs.Fields("ContractNumber")%></td>                   
                        <td height="5" width="86"><%= FormatDate(rs.Fields("LotTDate")) %></td>
                        <td  height="5"  width="377"><%=rs.Fields("OrdDetailSecurity")%></td>
                        <td align=right  height="5" width="90"><%= FormatNum(rs.Fields("LotQty")) %></td>
                        <td align=right   width="80" height="5">&nbsp;&nbsp;<%= FormatNum(rs.Fields("LotPrice")) %></td>                        
                        <td align=right width="91" height="5" >&nbsp;&nbsp;<%= FormatNum(rs.Fields("NetAmount")) %></td>
                        
                </tr>
                <%
             Rs.MoveNext   
             If Not (Rs.EOF Or Rs.BOF) Then
				If Rs.Fields("Client_DPA_").Value <> currClientCode Then					
					Rs.Move -1
					isDifferentClient = True
				End If
             Else
				isDifferentClient = True
				Rs.Move -1
             End If   
         Loop      
                %>
                
           <tr>
						<td colspan=7 align=right  height="51" width="735">&nbsp;</td>
           </tr>               
        
         <tr>
						<td colspan=6 align=right  height="31" width="644">
                        <font size="4" face="Courier">CLIENT TOTALS:</font></td>
                        <td align=right height="31" width="89">&nbsp;&nbsp;<%= FormatNum(totalNetAmount) %></td>
                        
         </tr>
         <tr>
		 <td width="43" height="16">
         <p></td>
		 </tr>
		 <tr>
		 <td width="43" height="16">
         <p></p>
         <p></p>
         <p></p>
         <p></p>
         <p></td>
		 </tr>
         <tr>
			<td colspan=5 width="476" height="22" ><b>
            <font face="Courier" size="4">
            Checked by : __________________________________________</font></b></td>      
			<td colspan=4 align=left width="259" height="22"><b>
            <font size="4" face="Courier">Cheque No:</font><font size="2" face="Courier">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
		 </tr>
		 <tr>
		 <td width="43" height="41"></td>
		 </tr>        
         <tr>
			<td colspan=4 width="476" height="22" ><b>
            <font face="Courier" size="4">
            Approved by :___________________________________________</font></b></td>      
			<td colspan=4 width="300" height="22" >&nbsp;</td>      
		 </tr>
</table>
<!--</center>-->
<%

			rs.MoveNext
			'important!		
			If Not Rs.EOF Then %>
				<BR class="newpage">
		<%	End If
        Loop
        
        conn.Close
        Set conn = Nothing
        Set Rs = Nothing%>


</body>