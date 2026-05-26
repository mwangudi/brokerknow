<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>CDS Bank Statement Account</title>

	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>



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
			margin-top: 1cm;    
			margin-bottom: 2cm;
			size: portrait;
			
			br.newpage{
				page-break-before:always;
			}
			
			
		}

	</style>

</head>

<body Class="Reports">
<!--#include file="../libroutines.asp"-->

<%

genReport = Request.Form("genReport")
selectedBank = 4
selectedFromDate = Request.Form("transFromDate")
selectedToDate = Request.Form("txtToDate")
SelectedType=5
FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)
thistype=Request.Form("Selectedtype")

If genReport <> "1" Or selectedBank = "" Or Not IsDate(selectedFromDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		
		function validateForm(frm){						
			frm.target = '_self';			
			frm.submit();
		}				
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdDate","<%= FormatDate(Date) %>",1);

	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="CDSStatement.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<% currentEntityType=5 %>
		<table border="0">
			<tr>
				<td>Select date from:</td>
				<td  colspan="2">
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>				
			</tr>
			<tr>
				<td>To date:</td>
				<td colspan="2">
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
				</td>
				
			</tr>			

			<tr>
				<td></td>
				<td><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp;</td>
				<td><input type = 'button' Class=Buttons name ='cmdRefresh' id = 'cmdRefresh' value=" Refresh " OnClick="JavaScript: window.location.replace('AccountsStatement.asp')"></td>				
				
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

'Response.write(selectedBank)
'Response.end

	Set conn = GetActiveConnection("KBroker")
	Set Rs = CreateObject("ADODB.Recordset")

	Select Case(SelectedType)
      	case 1						        
			sqlStr = "SELECT * FROM StatementList1 WHERE Client_DPA_ = " & selectedBank & " AND TransDate Between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "'"
            case 2
			sqlStr = "SELECT * FROM AgentStatement1 where Agent_DPA_ LIKE '" & selectedBank & "' AND TransDate Between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "'"
		case 3
			sqlStr = "SELECT * FROM BrokerStatement1 where Broker_DPA_ LIKE '" & selectedBank & "' AND TransDate Between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "'"
		case 4
			sqlStr = "SELECT * FROM BrokerCommissionStatement where Entity_DPA_ =" & selectedBank & " AND TransDate between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "' ORDER BY TransDate, EntityName DESC"						
		case 5
			sqlStr = "SELECT * FROM CDSSettlements where BankAccount_DPA_ =" & selectedBank & " AND TransDate between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "' ORDER BY TransDate, REF"		
		case 6
			sqlStr = "SELECT * FROM CommissionStatement where Entity_DPA_ =" & selectedBank & " AND TransDate between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "' and Balance <>0 order by TransDate asc"		
		case 7
			sqlStr = "SELECT * FROM OwnerStatement where Owner_DPA_ =" & selectedBank & " AND TransDate between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "'"		
		case 8
			sqlStr = "SELECT * FROM CDSControlStatement where Client_DPA_ =" & selectedBank & " AND TransDate between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "' order by Transdate,ref"		
		case 11
			sqlStr = "SELECT * FROM ComputerStatement where Entity_DPA_ =" & selectedBank & " AND TransDate between '" & FormatDate(CDate(selectedFromDate)) & "' and '" & FormatDate(CDate(selectedToDate)) & "'"				
      end select
	 
	'Response.write(sqlStr)
	'Response.end   

	Rs.CursorLocation = adUseClient	
	Rs.Open SQLServerFormat(HandleQuote(sqlStr)), conn.ConnectionString, adOpenKeyset, adLockOptimistic
	'Rs.Filter = ""
	
	If rs.EOF Or rs.BOF Then%>
		<Script Language="JavaScript">
			alert("The specified account does not have any transaction using the specified date criterion")
			window.history.go(-1);
		</Script>
		<%Set Rs = Nothing
		Set Conn = Nothing
		Response.End
	End If
	
	Set rsBank = Conn.Execute ("SELECT * FROM CompleteEntityList WHERE (EntityType_DPA_ = " & SelectedType & ") AND (Entity_DPA_ = " & selectedBank & ")")
	
	If Not (rsBank.EOF Or rsBank.BOF) Then
		accountDesc = rsBank.Fields("EntityName").Value
		'accountAddress = rsBank.Fields("ClientAddr").Value
	End If
	
	Set rsBank = Nothing
	Rs.PageSize=50

	Rs.CacheSize = Rs.PageSize
		intPageCount = Rs.PageCount 
		intRecordCount = Rs.RecordCount 
	
		first=0
		
		'Response.write(intPagecount)
		
		'Rs.Getrows(10)
		'="Select Top 10"
		
		
		
	' Now you must double check to make sure that you are not before the start
	' or beyond end of the recordset.  If you are beyond the end, set 
	' the current page equal to the last page of the recordset.  If you are
	' before the start, set the current page equal to the start of the recordset.	

	Rs.CacheSize = Rs.PageSize
	intPageCount = Rs.PageCount 
	intRecordCount = Rs.RecordCount 
		
        PageNumber1=PageNumber1 + 1
        
        intPage=0
        
         intPageCount=Cint(intPageCount)
      m=0  
      t=0   
	totalDebits = 0
    	totalCredits = 0
    
	RunningBal=0
	firsttime=0

	Do while Cint(intPage) < intPageCount

	intPage=intPage + 1	
	'Response.write(intpage)
	
	if(Cint(first)=1) then
	%>
             <BR class="newpage">
    <%
	end if
	

	If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
	If CInt(intPage) <= 0 Then intPage = 1
	
	 'Make sure that the recordset is not empty.  If it is not, then set the 
	 'AbsolutePage property and populate the intStart and the intFinish variables.
	
	'if Not(Rs.eof and Rs.bof) Then

	If intRecordCount > 0 Then 'and Not(Rs.eof and Rs.bof) 
		Rs.AbsolutePage = intPage
		intStart = Rs.AbsolutePosition
		'Response.write(intStart)
		
		If CInt(intPage) = CInt(intPageCount) Then
			intFinish = intRecordCount
		Else
			intFinish = intStart + (Rs.PageSize - 1)
		End if
	End If	  

	isOpeningBalance = CBool(Rs.Fields("IsOpeningBalance").Value)
	
    if first=0 then	
	If Not IsOpeningBalance Then
		'get latest prev balance
		
		Select Case(SelectedType)
      		case 1						        	
		sqlStr = "SELECT SUM(ISNULL(dbo.StatementList1.Credit - dbo.StatementList1.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal, dbo.StatementList1.Client_DPA_" & _
				 " FROM  dbo.StatementList1 INNER JOIN dbo.Client ON dbo.StatementList1.Client_DPA_ = dbo.Client.Client_DPA_" & _
				 " WHERE (dbo.Client.Deleted = 0) AND (dbo.StatementList1.TransDate < '" & FormatDate(selectedFromDate) & "') GROUP BY dbo.StatementList1.Client_DPA_, dbo.Client.ClientOpeningBal having dbo.StatementList1.Client_DPA_=" & selectedBank 
		case 2
		sqlStr = "SELECT SUM(ISNULL(dbo.AgentStatement1.Credit - dbo.AgentStatement1.Debit, 0)) + dbo.Agent.AgentOpeningBal AS CurrentBal, dbo.AgentStatement1.Agent_DPA_" & _
				 " FROM  dbo.AgentStatement1 INNER JOIN dbo.Agent ON dbo.AgentStatement1.Agent_DPA_ = dbo.Agent.Agent_DPA_" & _
				 " WHERE (dbo.Agent.Deleted = 0) AND (dbo.AgentStatement1.TransDate < '" & FormatDate(selectedFromDate) & "') GROUP BY dbo.AgentStatement1.Agent_DPA_, dbo.Agent.AgentOpeningBal having dbo.AgentStatement1.Agent_DPA_=" & selectedBank 				
		case 3
		
		sqlStr = "SELECT SUM(ISNULL(dbo.BrokerStatement1.Credit - dbo.BrokerStatement1.Debit, 0)) + dbo.Broker.BrokerOpeningBal AS CurrentBal, dbo.BrokerStatement1.Broker_DPA_" & _
				 " FROM  dbo.BrokerStatement1 INNER JOIN dbo.Broker ON dbo.BrokerStatement1.Broker_DPA_ = dbo.Broker.Broker_DPA_" & _
				 " WHERE (dbo.BrokerStatement1.TransDate < '" & FormatDate(selectedFromDate) & "') GROUP BY dbo.BrokerStatement1.Broker_DPA_, dbo.Broker.BrokerOpeningBal having dbo.BrokerStatement1.Broker_DPA_=" & selectedBank  
		case 4
		sqlStr="SELECT SUM(ISNULL(Balance, 0)) AS CurrentBal, Entity_DPA_ FROM  dbo.BrokerCommissionStatement" & _
			   " WHERE (TransDate < '" & FormatDate(selectedFromDate) & "') GROUP BY Entity_DPA_ having Entity_DPA_=" & selectedBank 						
		case 5
		sqlStr = "SELECT CASE (CDSSettlements.IsOpeningBalance) WHEN 0 THEN SUM(ISNULL(dbo.CDSSettlements.CreditBal - dbo.CDSSettlements.Debit, 0))" & _ 
                 " ELSE 0 END + dbo.Account.AccountOpeningBal AS CurrentBal, dbo.CDSSettlements.BankAccount_DPA_" & _
				 " FROM  dbo.CDSSettlements INNER JOIN dbo.Account ON dbo.CDSSettlements.BankAccount_DPA_ = dbo.Account.Account_DPA_" & _
				 " WHERE  (dbo.CDSSettlements.TransDate < '" & FormatDate(selectedFromDate) & "') GROUP BY dbo.CDSSettlements.BankAccount_DPA_,CDSSettlements.IsOpeningBalance, dbo.Account.AccountOpeningBal having dbo.CDSSettlements.BankAccount_DPA_=" & selectedBank  
		case 6
		sqlStr="SELECT SUM(ISNULL(Balance, 0)) AS CurrentBal, Entity_DPA_ FROM  dbo.CommissionStatement" & _
			   " WHERE (TransDate < '" & FormatDate(selectedFromDate) & "') GROUP BY Entity_DPA_ having Entity_DPA_=" & selectedBank 						
		case 7
		sqlStr = "SELECT SUM(ISNULL(dbo.OwnerStatement.Credit - dbo.OwnerStatement.Debit, 0)) + dbo.Owner.OwnerOpeningBal AS CurrentBal, dbo.OwnerStatement.Owner_DPA_" & _
				 " FROM  dbo.OwnerStatement INNER JOIN dbo.Owner ON dbo.OwnerStatement.Owner_DPA_ = dbo.Owner.Owner_DPA_" & _
				 " WHERE  (dbo.OwnerStatement.TransDate < '" & FormatDate(selectedFromDate) & "') GROUP BY dbo.OwnerStatement.Owner_DPA_, dbo.Owner.OwnerOpeningBal having dbo.OwnerStatement.Owner_DPA_=" & selectedBank 				
		case 8
		sqlStr="SELECT SUM(ISNULL(case(IsOpeningBalance) when 1 then Balance else Credit - Debit end , 0)) AS CurrentBal, Client_DPA_ FROM  dbo.CDSControlStatement" & _
			   " WHERE (TransDate < '" & FormatDate(selectedFromDate) & "') GROUP BY Client_DPA_"				
		
		case 11
		sqlStr="SELECT SUM(ISNULL(Credit - Debit, 0)) AS CurrentBal, Entity_DPA_ FROM  dbo.ComputerStatement" & _
			   " WHERE (TransDate < '" & FormatDate(selectedFromDate) & "') GROUP BY Entity_DPA_ having Entity_DPA_=" & selectedBank 						
		
		end select 	    
		
		
		'Response.write(sqlStr)
		'Response.end

		Set cloneRs = Conn.Execute(sqlStr)
		
		if not(cloneRs.eof and cloneRs.bof)	then
		OpeningBalance =  cloneRs.Fields("CurrentBal").Value		
		else
		OpeningBalance=0
		end if	
			
	Else
		OpeningBalance =  Rs.Fields("CreditBal").Value		
		
	End If
     end if	
	first=1
'Response.write(OpeningBalance)	
%>	


<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="10%" nowrap><font face="Impact" size="4">CDS STATEMENT</font></td>
      <td width="60%" nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
      
    </tr>

  </table>
<br>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
    <tr>
      <td width="1%"><b>Date:</b></td>
      <td width="48%"><%= FormatDate(Date) %></td>
    </tr>

    <tr>
      <td width="1%"><b>Account:</b></td>
      <td width="48%"><%= accountDesc %></td>
    </tr>


</table>
<BR>



  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="700">        
      <% if(SelectedType<>5) then
      %>
		<tr>
		<td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Date&nbsp;&nbsp;&nbsp;&nbsp;</font></b></td>
		<td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Ref:</font></b></td>      
		<td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Particulars:</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">Debit:</font></b></td>
		<td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align=right><b><font face="Arial Narrow" size="3">Credit:</font></b></td>
		<td align="right" style="border-right-style: solid; border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">Balance</font></b></td>
		</tr>
	  <%
	    if(m=0) then
	    If Not IsOpeningBalance Then %>
		<tr>	
		  <td><%= FormatDate(selectedFromDate) %></td>		  
		  <td>&nbsp;</td>
		  <td>Opening Balance</td>
		  <td align="right">&nbsp;</td>
		  <td align="right">&nbsp;</td>
		  <td align="right"><%= FormatNum(CreditDebitValue(OpeningBalance)) %></td>
		</tr>    
		<%	runningBal = CreditDebitValueRev(OpeningBalance)
			OpeningBalance = OpeningBalance
			Set cloneRs = Nothing
		m=1
		End If
		
		end if    
		else
		
       %>
      <tr>
      <td style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" width="70"><b><font face="Arial Narrow" size="3">Date</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Ref:</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Receipt No:</font></b></td>
      <td style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"><b><font face="Arial Narrow" size="3">Particulars:</font></b></td>
      <td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right><b><font face="Arial Narrow" size="3">Debit:</font></b></td>
      <td align="right" style="border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1" align=right><b><font face="Arial Narrow" size="3">Credit:</font></b></td>
      <td align="right" style="border-right-style: solid; border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1"  align=right width="100"><b><font face="Arial Narrow" size="3">Balance</font></b></td>
	  </tr>
    
	  <%
		  if(m=0) then
			  If Not IsOpeningBalance Then%>
				<tr>	
				  <td><%= FormatDate(selectedFromDate) %></td>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				  <td>Opening Balance</td>
				  <td align="right">&nbsp;</td>
				  <td align="right">&nbsp;</td>
				  <td align="right"><%= FormatNum(CreditDebitValue(OpeningBalance)) %></td>
				</tr>    
			<%
				if(firsttime=0) then
				runningBal = CreditDebitValueRev(OpeningBalance)
				end if

				OpeningBalance = OpeningBalance
				Set cloneRs = Nothing
				m=1
			End If
		end if
	end if    
    
    ''Response.write(OpeningBalance)
    ''Response.end
    
    'Do Until Rs.EOF           
	'Response.write(runningBal)
	
	firsttime=1

    For intRecord = 1 to Rs.PageSize					
		%>
		<tr>	
		  <td><%= FormatDate(Rs.Fields("TransDate").Value) %></td>
		  <td><%= Rs.Fields("Ref").Value %></td>
		  <% 
		  if(SelectedType=1) then		  		  
		   else
		   if(SelectedType<>5) then		   
			else
				if(Rs.Fields("ReceiptNo").Value =0 ) then %>
				<td>&nbsp;</td>
				<% else %>
				<td><%= Rs.Fields("ReceiptNo").Value %></td>
				<% end if
			end if
				end if %>
		  <td><%= Rs.Fields("Particulars").Value %></td>
		  <td align="right"><% 
							   if(rs("CDS")=1) then							   
									if(Cdbl(Rs.Fields("CreditBal").Value)>0) then
									Response.Write FormatNum(Abs(Rs.Fields("CreditBal").Value)) 
									totalDebits = totalDebits + Rs.Fields("CreditBal").Value 		
									%>
									</td><td>&nbsp;</td>
									<% else %>
									<td align="right"><%=FormatNum(Abs(Rs("CreditBal")))%></td>
									<%
									totalCredits = totalCredits + Abs(Rs.Fields("CreditBal").Value)
									end if
							   else
							   If Rs.Fields("Debit").Value <> "0" Then
									totalDebits = totalDebits + Rs.Fields("Debit").Value 								        
									Response.Write FormatNum(Rs.Fields("Debit").Value)
							   End If %>
			</td>
		  <td align="right"><%If Rs.Fields("Credit").Value Then
									totalCredits = totalCredits + Abs(Rs.Fields("Credit").Value)
									Response.Write FormatNum(Rs.Fields("Credit").Value) 
							  End If
						End if
							  %>
		  </td>
		  <td align="right">

			<%			
			If Not IsOpeningBalance Then					
					if(rs("CDS")=1) then					
					runningBal = runningBal - (Rs.Fields("CreditBal").Value)
					else
					runningBal = runningBal + (Rs.Fields("Credit").Value - Rs.Fields("Debit").Value)					
					end if
					Response.Write  FormatNum(CreditDebitValue(runningBal)) 
				Else

					if(SelectedType=1 or SelectedType=2 or SelectedType=8) then
						if(rs("IsOpeningBalance")=1) then						
							Response.Write FormatNum(CreditDebitValue(Rs.Fields("Balance").Value))						
							runningBal = rs("CreditBal")
						else
							Valuewrite=(Abs(Rs.Fields("Credit").Value) - Rs.Fields("Debit").Value)	
							
							Response.write(FormatNum(CreditDebitValue(runningBal+ValueWrite)))											
							runningBal = runningBal + Valuewrite
							'Response.Write FormatNum(Rs.Fields("Balance").Value)										
						end if
      		        else
					''Response.write(rs("IsOpeningBalance"))
					
					if(SelectedType=7) then
					runningBal = runningBal + (Rs.Fields("Credit").Value-Rs.Fields("Debit").Value)					
					else
						if(rs("CDS")=1) then
						runningBal = runningBal - Rs.Fields("CreditBal").Value					
						else
						'Response.write(runningBal)
							if(rs("IsOpeningBalance")=1) then						
							runningBal = runningBal + Rs.Fields("CreditBal").Value				
							else
							runningBal = runningBal + (Rs.Fields("Credit").Value-Rs.Fields("Debit").Value)					
							end if
						end if
					end if
					
					Response.Write FormatNum(CreditDebitValue(runningBal))
					end if
				End If 				
				%>
		  </td>
		</tr>
	
	<%	
		'end if
		'Rs.MoveNext	
	lastDate = Rs.Fields("TransDate").Value
	Total=Total + RunningBal
	rs.MoveNext
        
		If Rs.EOF Then Exit for

        Next
'Response.write(lastDate)
'Response.end

if Cint(intPage) <> intPageCount then
	if(SelectedType<>5) then 	
  		%>
  		<tr>
      	<td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        	&nbsp;&nbsp;&nbsp; </td>
    		</tr>

    		<tr>
      	<td colspan="5" align="right">
        	<font size="1">
        	Opening Balance:</font></td>
      	<td align="right"><font size="1"><%= FormatNum(CreditDebitValue(OpeningBalance)) %></font>&nbsp;</td>
    		</tr>

    		<tr>
      	<td colspan="5" align="right"><font size="1">less Running Debits:</font></td>

      	<td align="right"><font size="1"><%= FormatNum(CreditDebitValue(0 - totalDebits)) %></font>&nbsp;</td>
    		</tr>

    		<tr>
      	<td colspan="5" align="right"><font size="1">add Running Credits:</font></td>

      	<td align="right"> <font size="1"> <%= FormatNum(CreditDebitValue(totalCredits)) %></font>&nbsp;</td>
    		</tr>

    		<tr>
      	<td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>

    		</tr>

    		<tr>
      	<td colspan="4">&nbsp;</td>
      	<td align="right"><b><font size="1">Running Balance:</font></b></td>
      	<td align="right"><b><font size="1"><%= CreditDebitValue(FormatNum(runningBal)) %></font></b>&nbsp;</td>
    		</tr>
   		<%
		else
		%>
  		<tr>
      	<td colspan="7" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        	&nbsp;&nbsp;&nbsp; </td>
    		</tr>

    		<tr>
      	<td colspan="6" align="right">
        	<font size="1">
        	Opening Balance:</font></td>
      	<td align="right"><font size="1"><%= FormatNum(CreditDebitValue(OpeningBalance)) %></font>&nbsp;</td>
    		</tr>

    		<tr>
      	<td colspan="6" align="right"><font size="1">less Running Debits:</font></td>

      	<td align="right"><font size="1"><%= FormatNum(CreditDebitValue(0 - totalDebits)) %></font>&nbsp;</td>
    		</tr>

    		<tr>
      	<td colspan="6" align="right"><font size="1">add Running Credits:</font></td>

      	<td align="right"> <font size="1"> <%= FormatNum(CreditDebitValue(totalCredits)) %></font>&nbsp;</td>
    		</tr>

    		<tr>
      	<td colspan="7" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>

    		</tr>

    		<tr>
      	<td colspan="5">&nbsp;</td>
      	<td align="right"><b><font size="1">Running Balance:</font></b></td>
      	<td align="right"><b><font size="1"><%= CreditDebitValue(FormatNum(runningBal)) %></font></b>&nbsp;</td>
    		</tr>
   		<%		
		end if
   	end if
   	%>	
	</table>
	<%	
	loop	
	%>
	<table border="0" cellspacing="1" cellpadding="0" style="font-family: Arial Narrow; LEFT-MARGIN:100PX"  width="700">     
<%
	 if(SelectedType<>5) then
	 %>
	
		<tr>	
		  <td><%= FormatDate(lastDate) %></td>		  
		  <td>&nbsp;</td>
		  <td>Closing Balance</td>
		  <td align="right">&nbsp;</td>
		  <td align="right">&nbsp;</td>
		  <td align="right"><%= CreditDebitValue(FormatNum(runningBal)) %></td>
		</tr>  
	
    <tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp; </td>

    </tr>

    <tr>
      <td colspan="5" align="right">
        Opening Balance:</td>

      <td align="right"><%= FormatNum(CreditDebitValue(OpeningBalance)) %></td>
    </tr>

    <tr>
      <td colspan="5" align="right">less Total Debits:</td>

      <td align="right"><%= FormatNum(CreditDebitValue(0 - totalDebits)) %></td>
    </tr>

    <tr>
      <td colspan="5" align="right">add Total Credits:</td>

      <td align="right"> <%= FormatNum(CreditDebitValue(totalCredits)) %></td>
    </tr>

    <tr>
      <td colspan="6" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>

    </tr>

    <tr>
      <td colspan="4"><font size="1">Current: <%= CreditDebitValue(FormatNum(runningBal)) %>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;30-60 Days:&nbsp;<%=CreditDebitValue(FormatNum(ClientBalance(30)))%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Over 60 Days:&nbsp;<%=CreditDebitValue(FormatNum(ClientBalance(60)))%></font> </td>
      <td align="right"><b><font size="1">Total Balance:</font></b></td>
      <td align="right"><b><font size="1"><%= CreditDebitValue(FormatNum(runningBal)) %></font></b>&nbsp;</td>
    </tr>

   <%
	 else
      %>
	
		<tr>	
		  <td><%= FormatDate(lastDate) %></td>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		  <td>Closing Balance</td>
		  <td align="right">&nbsp;</td>
		  <td align="right">&nbsp;</td>
		  <td align="right"><%= CreditDebitValue(FormatNum(runningBal)) %></td>
		</tr>  
	
    <tr>
      <td colspan="7" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp; </td>

    </tr>

    <tr>
      <td colspan="6" align="right">
        Opening Balance:</td>

      <td align="right"><%= FormatNum(CreditDebitValue(OpeningBalance)) %></td>
    </tr>

    <tr>
      <td colspan="6" align="right">less Total Debits:</td>

      <td align="right"><%= FormatNum(CreditDebitValue(0 - totalDebits)) %></td>
    </tr>

    <tr>
      <td colspan="6" align="right">add Total Credits:</td>

      <td align="right"> <%= FormatNum(CreditDebitValue(totalCredits)) %></td>
    </tr>

    <tr>
      <td colspan="7" align="right" style="border-bottom-style: solid; border-bottom-width: 1">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>

    </tr>

    <tr>
      <td colspan="5"><font size="1">Current: <%= CreditDebitValue(FormatNum(runningBal)) %>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;30-60 Days:&nbsp;<%=CreditDebitValue(FormatNum(ClientBalance(30)))%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Over 60 Days:&nbsp;<%=CreditDebitValue(FormatNum(ClientBalance(60)))%></font> </td>
      <td align="right"><b><font size="1">Total Balance:</font></b></td>
      <td align="right"><b><font size="1"><%= CreditDebitValue(FormatNum(runningBal)) %></font></b>&nbsp;</td>
    </tr>
    
   <% 
   end if
   
   function ClientBalance(Days)
   
	Select Case(SelectedType)
      	case 1						        	
		sqlStr = "SELECT SUM(ISNULL(dbo.StatementList1.Credit - dbo.StatementList1.Debit, 0)) + dbo.Client.ClientOpeningBal AS CurrentBal, dbo.StatementList1.Client_DPA_" & _
				 " FROM  dbo.StatementList1 INNER JOIN dbo.Client ON dbo.StatementList1.Client_DPA_ = dbo.Client.Client_DPA_" & _
				 " WHERE (dbo.Client.Deleted = 0) AND (dbo.StatementList1.TransDate < '" & FormatDate(CDate(selectedToDate)-Days) & "') GROUP BY dbo.StatementList1.Client_DPA_, dbo.Client.ClientOpeningBal having dbo.StatementList1.Client_DPA_=" & selectedBank 
		case 2
		sqlStr = "SELECT SUM(ISNULL(dbo.AgentStatement1.Credit - dbo.AgentStatement1.Debit, 0)) + dbo.Agent.AgentOpeningBal AS CurrentBal, dbo.AgentStatement1.Agent_DPA_" & _
				 " FROM  dbo.AgentStatement1 INNER JOIN dbo.Agent ON dbo.AgentStatement1.Agent_DPA_ = dbo.Agent.Agent_DPA_" & _
				 " WHERE (dbo.Agent.Deleted = 0) AND (dbo.AgentStatement1.TransDate < '" & FormatDate(CDate(selectedToDate)-Days) & "') GROUP BY dbo.AgentStatement1.Agent_DPA_, dbo.Agent.AgentOpeningBal having dbo.AgentStatement1.Agent_DPA_=" & selectedBank 				
		case 3
		
		sqlStr = "SELECT SUM(ISNULL(dbo.BrokerStatement1.Credit - dbo.BrokerStatement1.Debit, 0)) + dbo.Broker.BrokerOpeningBal AS CurrentBal, dbo.BrokerStatement1.Broker_DPA_" & _
				 " FROM  dbo.BrokerStatement1 INNER JOIN dbo.Broker ON dbo.BrokerStatement1.Broker_DPA_ = dbo.Broker.Broker_DPA_" & _
				 " WHERE (dbo.BrokerStatement1.TransDate < '" & FormatDate(CDate(selectedToDate)-Days) & "') GROUP BY dbo.BrokerStatement1.Broker_DPA_, dbo.Broker.BrokerOpeningBal having dbo.BrokerStatement1.Broker_DPA_=" & selectedBank  
		case 4
		sqlStr="SELECT SUM(ISNULL(Credit - Debit, 0)) AS CurrentBal, Entity_DPA_ FROM  dbo.BrokerCommissionStatement" & _
			   " WHERE (TransDate < '" & FormatDate(CDate(selectedToDate)-Days) & "') GROUP BY Entity_DPA_ having Entity_DPA_=" & selectedBank 						
		case 5
		sqlStr = "SELECT SUM(ISNULL(dbo.CDSSettlements.CreditBal - dbo.CDSSettlements.Debit, 0)) + dbo.Account.AccountOpeningBal AS CurrentBal, dbo.CDSSettlements.BankAccount_DPA_" & _
				 " FROM  dbo.CDSSettlements INNER JOIN dbo.Account ON dbo.CDSSettlements.BankAccount_DPA_ = dbo.Account.Account_DPA_" & _
				 " WHERE  (dbo.CDSSettlements.TransDate < '" & FormatDate(CDate(selectedToDate)-Days) & "') GROUP BY dbo.CDSSettlements.BankAccount_DPA_, dbo.Account.AccountOpeningBal having dbo.CDSSettlements.BankAccount_DPA_=" & selectedBank  
		case 6
		sqlStr="SELECT SUM(ISNULL(Credit - Debit, 0)) AS CurrentBal, Entity_DPA_ FROM  dbo.CommissionStatement" & _
			   " WHERE (TransDate < '" & FormatDate(CDate(selectedToDate)-Days) & "') GROUP BY Entity_DPA_ having Entity_DPA_=" & selectedBank 						
		case 7
		sqlStr = "SELECT SUM(ISNULL(dbo.OwnerStatement.Credit - dbo.OwnerStatement.Debit, 0)) + dbo.Owner.OwnerOpeningBal AS CurrentBal, dbo.OwnerStatement.Owner_DPA_" & _
				 " FROM  dbo.OwnerStatement INNER JOIN dbo.Owner ON dbo.OwnerStatement.Owner_DPA_ = dbo.Owner.Owner_DPA_" & _
				 " WHERE (dbo.OwnerStatement.TransDate < '" & FormatDate(CDate(selectedToDate)-Days) & "') GROUP BY dbo.OwnerStatement.Owner_DPA_, dbo.Owner.OwnerOpeningBal having dbo.OwnerStatement.Owner_DPA_=" & selectedBank 				
		case 8
		sqlStr="SELECT SUM(ISNULL(Credit - Debit, 0)) AS CurrentBal, Client_DPA_ FROM  dbo.CDSControlStatement" & _
			   " WHERE (TransDate < '" & FormatDate(CDate(selectedToDate)-Days) & "') GROUP BY Client_DPA_"		
		case 11
		sqlStr="SELECT SUM(ISNULL(Credit - Debit, 0)) AS CurrentBal, Entity_DPA_ FROM  dbo.ComputerStatement" & _
			   " WHERE (TransDate < '" & FormatDate(CDate(selectedToDate)-Days) & "') GROUP BY Entity_DPA_ having Entity_DPA_=" & selectedBank 						
			   		
		end select 		
		

Set Rs = Conn.Execute(sqlStr)
 if not(Rs.eof and Rs.Bof) then
 ClientBalance=Rs("CurrentBal")
 else
 ClientBalance=0
 end if
end function 


   %>
  </table>
<form method="POST" action="AccountsStatement.asp" Name="frmMain" id="frmMain">
<input type = 'hidden' name ='Selectedtype' id = 'Selectedtype' value='<%=Selectedtype%>'>   
</form>
</body>
</html>