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
			
			margin-left: 2cm;
			margin-right: 2cm;
			margin-top: 2cm;    
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
		
			SettlementDate = LTdate(CDate(rs("SettlementDate")),NoOfDays)	
			
			SettlementDate = FormatDate(SettlementDate)
			
			
			Conn.Execute("Insert Into PaymentVoucherSettlementDates (Contract_DPA_,SettlementDate) Values( " & rs("Contract_DPA_") & ",'" & SettlementDate & "')")
   			rs.movenext
   			loop
   		end if     
		
		set rs=nothing
		
		sqlStr = "SELECT * FROM [SettlePaymentVoucher] WHERE Day(SettlementDate) = Day('" & FormatDate(selectedTradeDate) & "') and " & _
                     " Month(SettlementDate) = Month('" & FormatDate(selectedTradeDate) & "') and Year(SettlementDate) = Year('" & FormatDate(selectedTradeDate) & "') ORDER BY ClientName"		
		
        
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
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
      <tr>
        <td width="100%" colspan="3" align="center">
        <b><font size="3">PAYMENT VOUCHER</font></b></td>
      </tr>
      <tr>
        <td width="60%">&nbsp;</td>
        <td width="40%" colspan="2">&nbsp;</td>
      </tr>
      <tr>
        <td width="60%" height="15">&nbsp;</td>
        <td width="40%" colspan="2" height="15">AFRICAN ALLIANCE Malawi SECURITIES</td>
      </tr>
      <tr>
        <td width="60%" height="15">&nbsp;</td>
        <td width="40%" colspan="2" height="15">(Members of the NSE)</td>
      </tr>
      <tr>
        <td width="60%" height="15" height="15">&nbsp;</td>
        <td width="40%" colspan="2" height="15">Ground Floor, Malawi Re Towers</td>
      </tr>
      <tr>
        <td width="60%" height="15">&nbsp;</td>
        <td width="40%" colspan="2" height="15">P.O.Box&nbsp; 27639 -&nbsp; 00506, NAIROBI 
        Malawi</td>
      </tr>
      <tr>
        <td width="60%" height="15">&nbsp;</td>
        <td width="40%" colspan="2" height="15">Tel:&nbsp; +254&nbsp; 020&nbsp; 2735138</td>
      </tr>
      <tr>
        <td width="60%" height="15">&nbsp;</td>
        <td width="40%" colspan="2" height="15">Fax: +254 020&nbsp;&nbsp; 2731162</td>
      </tr>
      <tr>
        <td colspan="3">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3">TO THE ACCOUNTANT</td>
      </tr>
      <tr>
        <td colspan="3">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3">PLEASE PAY</td>
      </tr>
      <tr>
        <td colspan="3">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3"><b><%= Rs.Fields("ClientName").Value %> &nbsp;[<%= Trim(Rs.Fields("Client_DPA_").Value)%>]</b>&nbsp;</td>
      </tr>
     
      <tr>
        <td colspan="3"><%= RsNames.Fields("ClientAddr").Value %>&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3">&nbsp;</td>
      </tr>
      <tr>
        <td width="60%">&nbsp;</td>
        <td width="40%" colspan="2">
		        <Table border="0" cellpadding="0" cellspacing="0" width="100%">
			         <tr>
			           <td width="50%">&nbsp;</td>
			           <td width="50%" colspan="2"><b>Cheque No:</b></td>
			          </tr>
		        </Table>
        </td>
      </tr>
      <tr>
        <td width="60%">&nbsp;</td>
        <td width="40%" colspan="2">
		        <Table border="0" cellpadding="0" cellspacing="0" width="100%">
			         <tr>
			           <td width="50%" >&nbsp;</td>
			           <td width="50%" colspan="2"><b>Payment Date:</b> &nbsp;<%= FormatDate(rs("SettlementDate")) %></td>
			         </tr>
		        </Table>
        </td>
      </tr>

      <tr>
        <td colspan="3">&nbsp;</td>
      </tr>
      <tr>
        <td width="66%" colspan="2">Being Payment For: </td>
        <td width="34%">&nbsp;</td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td width="100%">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3">
      <tr bgcolor="#000000">
        <td width="5%" height="10"><b><FONT color="#ffffff">&nbsp;Slip</FONT></b></td>
        <td width="7%" height="10"><b><FONT color="#ffffff">&nbsp;Contract</FONT></b></td>
        <td width="14%" height="10"><b><FONT color="#ffffff">&nbsp;Traded</FONT></b></td>
        <td width="34%" height="10"><b><FONT color="#ffffff">&nbsp;Security</FONT></b></td>
        <td width="15%" height="10" align="right"><b><FONT color="#ffffff">&nbsp;Quantity</FONT></b>&nbsp;</td>
        <td width="10%" height="10" align="right"><b><FONT color="#ffffff">&nbsp;Price</FONT></b>&nbsp;</td>
        <td width="15%" height="10" align="right"><b><FONT color="#ffffff">Net&nbsp;Amount</FONT></b>&nbsp;</td>
      </tr>
      <%		
	totalNetAmount = 0
	totalQuantity = 0
	isDifferentClient = False
	
		 Do Until isDifferentClient 
			totalNetAmount = totalNetAmount + FormatNum(rs.Fields("NetAmount")) 
			totalQuantity = totalQuantity + FormatNum(rs.Fields("LotQty"))
			%>
                <tr>
                        <td  height="5" width="5%">&nbsp;<%=rs.Fields("LotSlipNo")%></td>
                        <td  height="5" width="7%">&nbsp;<%=rs.Fields("ContractNumber")%></td>                   
                        <td height="5" width="14%">&nbsp;<%= FormatDate(rs.Fields("LotTDate")) %></td>
                        <td  height="5"  width="34%">&nbsp;<%=rs.Fields("OrdDetailSecurity")%></td>
                        <td align="right"  height="5" width="15%"><%= FormatNum(rs.Fields("LotQty")) %>&nbsp;</td>
                        <td align="right"   width="10%" height="5">&nbsp;&nbsp;<%= FormatNum(rs.Fields("LotPrice")) %>&nbsp;</td>                        
                        <td align="right" width="15%" height="5" >&nbsp;&nbsp;<%= FormatNum(rs.Fields("NetAmount")) %>&nbsp;</td>
                        
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
        <td colspan="6">&nbsp;</td>
      </tr>
      <tr>
        <td width="65%" colspan="4" align="right"><b>Client Totals&nbsp;
        <td align="right" width="15%" align="right" style="border-bottom-style: solid; border-bottom-width: 1; border-top-style: solid; border-top-width: 1"><%= FormatNum(totalQuantity) %>&nbsp;</td>
        <td align="right" width="15%" align="right" style="border-bottom-style: solid; border-bottom-width: 1; border-top-style: solid; border-top-width: 1">&nbsp;</td>
        <td align="right" width="15%" align="right" style="border-bottom-style: solid; border-bottom-width: 1; border-top-style: solid; border-top-width: 1"><%= FormatNum(totalNetAmount) %>&nbsp;</td>
        </b></td>

      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td width="100%">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%">Checked by:&nbsp;  _____________________________ </td>
  </tr>
  <tr>
    <td width="100%">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%">Approved By:&nbsp; _____________________________</td>
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