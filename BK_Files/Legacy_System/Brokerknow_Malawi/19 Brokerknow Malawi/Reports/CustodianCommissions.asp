<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Contract Schedule</title>
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
			
			margin-left: 1cm;
			margin-right: 1cm;
			margin-top: 1cm;    
			margin-bottom: 1cm;
			writing-mode: tb-rl;
			height: 90%;
			margin: 10% 0%;			
			
			br.newpage{
				page-break-before:always;
			}
			
			
		}
		 
		
	</style>

<script language="javascript">

function ToggleClient (ctrl){

if (ctrl.checked==true){
 document.getElementById("cboCustodian").style.display = "";
} else{
  document.getElementById("cboCustodian").style.display = "none";
}

}
</script>

</head>

<body Class="Reports">

<Script Language="JavaScript">
	function HideRemindSelectLandscape(){
		try{			
			document.getElementById('landRem').style.display = 'none';
		}	
		catch(e){}	
	}
	function ShowRemindSelectLandscape(){
		try{			
			document.getElementById('landRem').style.display = '';
		}	
		catch(e){}	
	}
	window.onbeforeprint = HideRemindSelectLandscape;
	window.onafterprint = ShowRemindSelectLandscape;
</Script>

<!--#include file="../libroutinesTEST.asp"-->

<%

const beginLeviesCol = 17

genReport = trim(Request.Form("genReport"))
selectedFromDate = trim(Request.Form("txtDate"))
selectedToDate = trim(Request.Form("txtTo"))
filtercustodian = trim(Request.Form("filterByCustodian"))
CustodianID = trim(Request.Form("cboCustodian"))

If genReport <> "1" Or Not IsDate(selectedFromDate) Then%>
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
		var calto=new ctlSpiffyCalendarBox("calto", "frmMain", "txtTo","cmdTo","<%= FormatDate(Date) %>",1);
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="CustodianCommissions.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>
			<tr>
				<td>From Date</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>To Date</td>
				<td>
					<SCRIPT language="JavaScript">calto.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>&nbsp;<input value='1' type="checkbox" name="filterByCustodian" id="filterByCustodian" onclick="javascript: ToggleClient(this);">
				<label for="filterByCustodian">Filter by Custodian </label></td>
				<td>
				<select style="width: 350px;" name = 'cboCustodian' id = "cboCustodian" size="1" style="display: none;">
				<option value = '0'>&nbsp;</option>
					<%
					Set conn = GetActiveConnection("KBroker")
					        
				sqlStr = " SELECT GenericSetting.GenericSetting_DPA_, GenericSetting.GenericSettingDescription " & _
							" FROM GenericSetting INNER JOIN " & _
							" Generic ON GenericSetting.Generic_DPA_ = Generic.Generic_DPA_ " & _
							" WHERE (GenericSetting.EntityType_DPA_ = 1) AND (Generic.Generic_DPA_ = 1) " & _
							" ORDER BY GenericSetting.GenericSettingDescription"

				


					Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                     
                     intrscount = rs.RecordCount
                     
                     if intrscount > 0 then
						rs.MoveFirst
                      
						rsdata = rs.getrows()
                      
						for intcount = 0 to intrscount-1
						  %>					                        
						    <option value = '<%=trim(rsdata(0,intcount))%>'><%=trim(rsdata(1,intcount))%></option>
						  <%
						next
                      
                     end if
					%>
					</select>
				</td>
			</tr>
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... " id=Button1 name=Button1>&nbsp;&nbsp;</td>
			</tr>
		</table>
		
	</form>
	
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>


<% DrawPageFunctions True, True, True, True

if trim(selectedToDate) <> "" then
 if formatdate(selectedFromDate) = formatdate(selectedToDate) then
   headerDescription = " for Deals traded on: " & formatdate(selectedFromDate) 
 else
   headerDescription = " for Deals traded between: " & formatdate(selectedFromDate) & " & " & formatdate(selectedToDate)
 end if
else
 headerDescription = " for Deals traded between: " & formatdate(selectedFromDate) & " & " & formatdate(Date())
end if

if filtercustodian = "" then filtercustodian = 0
if selectedToDate = "" then selectedToDate = formatdate(now())
		
%>

<p id="toPDFOrient" name="toPDFOrient" value="L" style="display:none;">&nbsp;
<p id="toPDF" name="toPDF">

<i id="landRem">Remember to select landscape settings while printing.</i>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4">CUSTODIAN COMMISSIONS</font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
    <tr>
	   <td COLSPAN=2><font face="Arial" size="2"> <%= headerDescription %></font></td>
	</tr>
    <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>				

    <table border="0" width="100%" cellPadding="2" cellSpacing=0>
    <tr bgColor="#000000">
			
	  <td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">Traded</font></b></td>
      <td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">Code</font></b></td>
      <td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">Client</font></b></td>
      <td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">Secu</font></b></td>
      <td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">Br</font></b></td>
      <td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">Contr</font></b></td>
      <td bgColor="#000000" nowrap align="left"><b><font color="#FFFFFF">CDS</font></b></td>
      <td bgColor="#000000" nowrap align="right"><b><font color="#FFFFFF">Price</font></b></td>
      <td bgColor="#000000" nowrap align="right"><b><font color="#FFFFFF">Qty</font></b></td>
      <!--<td bgColor="#000000" nowrap>&nbsp;</td>-->
      <%
		Dim fld
		Dim conn 
		Dim sqlStr
		Dim rs
		Dim i
		Dim dailyTotalsArray()
			
		Set conn = GetActiveConnection("KBroker")
		
		'Set date filters
		
		if selectedToDate = "" then
		 selectedToDate = formatdate(Date())
		end if
		
		selectedFromDate = formatdate(selectedFromDate)
		selectedToDate = formatdate(selectedToDate)
		
 		sqlStr = "CustodianContractLevies_v2 '" & selectedFromDate & "','" & selectedToDate & "'"
		
		set rs = conn.execute(sqlStr)
		
		if filtercustodian = 1 then
		 Rs.Filter = "GenericSetting_DPA_ =" & CustodianID
		end if
		
		'i = 0
		'fldCount = fldCount + 2 'this is hard coding just to skip some unwanted columns 
		'From Patrick:  Do not include CDSC levy
		'               Include half of Compensation Levy
		
		for i = beginLeviesCol to rs.fields.count - 1
		    if i <> beginLeviesCol-1 then
				Redim Preserve dailyTotalsArray(i - beginLeviesCol)
				dailyTotalsArray(i - beginLeviesCol) = 0
				%>
					<td bgColor="#000000" nowrap  align="right"><b><font color="#FFFFFF">
					<%
					'Skip CDSC column
						If (Len(rs.fields(i).name)>5) Then
							Response.Write Mid(rs.fields(i).name,1,5)
						Else
							Response.Write rs.fields(i).name
						End If
					%></font></b></td>		
				<%
				'i = i + 1
			End if
		next
		
		'add gross, and net amount
		Redim Preserve dailyTotalsArray((i) - beginLeviesCol)
		dailyTotalsArray((i) - beginLeviesCol) = 0
		
		Redim Preserve dailyTotalsArray((i + 1) - beginLeviesCol)
		dailyTotalsArray((i + 1) - beginLeviesCol) = 0
		
		'Redim Preserve dailyTotalsArray((i + 2) - beginLeviesCol)
		'dailyTotalsArray((i + 2) - beginLeviesCol) = 0
      %>
		
		<!--<td bgColor="#000000" nowrap align="right"><b><font color="#FFFFFF">Gross</font></b></td>-->
		<!--<td bgColor="#000000" nowrap align="right"><b><font color="#FFFFFF">Net</font></b></td>-->
    </tr>
 <%
     If Not(rs.EOF Or rs.BOF) Then
        'rs.MoveFirst
        Do Until rs.EOF%>
        		<tr>
      <td nowrap align="left"><font size="1"><%= Day(rs.Fields("LotTDate")) & " " & MonthName(Month(rs.Fields("LotTDate")), True) %></font></td>
      <td nowrap align="left"><font size="1"><%= rs.Fields("Client_DPA_") %></font></td>
      <td nowrap align="left"><font size="1">
	  							<% If Len(rs.Fields("ClientName")) > 25 Then 
									Response.Write "..." &  Right(rs.Fields("ClientName"), 25)
								   Else
									Response.Write rs.Fields("ClientName")	
								   End If	 %></font></td>
      <td nowrap align="left"><font size="1"><%=rs.Fields("SecurityCode")%></font></td>
      <td nowrap align="left"><font size="1"><%=rs.Fields("BrokerCode")%></font></td>
      <td nowrap align="left"><font size="1"><%=rs.Fields("ContractNumber")%></font></td>
      <td nowrap align="left"><font size="1"><%=rs.Fields("LotSlipNo")%></font></td>
      <td nowrap align="right"><font size="1"><%=FormatNum(rs.Fields("LotPrice"))%></font></td>
      <td style="border-right:black 1px inset;background-color:transparent;" nowrap align="right"><font size="1"><%=FormatNumCommasOnly(rs.Fields("LotQty"))%></font></td>       
      <!--<td style="BORDER-RIGHT: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent">&nbsp;</td>-->
           <%	levyTotals = 0
		   'response.write beginLeviesCo & "   :  " & rs.fields.count - 1 :response.end
				for i = beginLeviesCol to rs.fields.count - 1
				    if i <> beginLeviesCol-1 then
				    
					  'Skip CDSC column
					  	If i = rs.Fields.Count - 2 Then
					  		myStyle =  "BORDER-RIGHT: black 1px inset; BACKGROUND-COLOR: transparent"
					  	Else
					  		myStyle = ""	
					  	End If
 	
					  	'Include only half of the compensation fund
					  	 'response.write rs.fields(i).value :response.end
					  	 if i = beginLeviesCol+7 then 
					  	  fieldvalue = cdbl(rs.fields(i).value/2)
					  	  fieldvalue = FormatNum(fieldvalue) ' compensation fund
					  	 else
					  	  fieldvalue = FormatNum(rs.fields(i).value)
					  	 end if
					  	 'response.write rs.fields(i).value :response.end
						 totalvalue = cdbl(rs.fields(17).value)   +  cdbl(rs.fields(18).value) +  cdbl(rs.fields(19).value) +  cdbl(rs.fields(20).value) +  cdbl(rs.fields(21).value) '+  cdbl(rs.fields(22).value) +  cdbl(rs.fields(23).value)
						 if i = 22 then
						 fieldvalue = FormatNum(totalvalue)  
						 end if
					  	%>
					  		<td  nowrap align=right  Style="<%= myStyle %>"><font size="1"><%=fieldvalue%></font></td>		
					  	<%
					  	'i = i + 1				
					  	levyTotals = levyTotals + fieldvalue	
					  	dailyTotalsArray(i - beginLeviesCol) = dailyTotalsArray(i - beginLeviesCol) + fieldvalue	
					
					end if  
				next
				
				'check whether this contract has the
				'levy of agent commission type, which is not
				'really a levy, but a rate of "broker commission" levy
				'if so, minus this value from the levyTotals variable
				sqlStr = "SELECT * FROM LevyContract WHERE Contract_DPA_ = " & rs.Fields("Contract_DPA_").Value & " AND SystemMaintained = 12"  
				Set tmpRs = Conn.Execute (SQLServerFormat(HandleQuote(sqlStr)))
				If Not (tmpRs.EOF OR tmpRs.BOF) Then
					'agent commission exists
					levyTotals = levyTotals - tmpRs.Fields("LevyAmount").Value
				End If
				Set tmpRs = Nothing		
				
				
				grossAmt = rs.Fields("LotGrossAmount") 'rs.Fields("LotPrice") * rs.Fields("LotQty") 
				
				If rs.Fields("OrderTypeSale").Value = 0 Then 
					netAmt = grossAmt + levyTotals 
				Else
					netAmt = grossAmt - levyTotals 
				End If
				
				dailyTotalsArray((i) - beginLeviesCol) = dailyTotalsArray((i) - beginLeviesCol) + grossAmt
				dailyTotalsArray((i + 1) - beginLeviesCol) = dailyTotalsArray((i + 1) - beginLeviesCol) + netAmt

				%>
             
             
             <!--<td nowrap align="right"><font size="1"><%'= FormatNum(grossAmt) %></font></td>       
             <td nowrap align="right"><font size="1"><%'= FormatNum(netAmt) %></font></td>--> 
             </tr>      
             <%rs.MoveNext
        Loop%>
        
        <tr>
			<td colspan="<%= 8 + UBound(dailyTotalsArray)%>">&nbsp;</td>
        </tr>
      
		<tr height="30px">
			<td colspan=1 align=right>
			</td>
			<td colspan=8 align=right>
				<b>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Daily totals:
				</b>	
			</td>	
			 <%	
				for i = 0 to 4'UBound(dailyTotalsArray)
						If i = 0 Then 
							myStyle = "BORDER-LEFT: #C0C0C0 1px inset; BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent; "
						ElseIf (i <> 0 And i <> UBound(dailyTotalsArray)) Then
							myStyle = "BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent;"
						Else
							myStyle = "BORDER-RIGHT: #C0C0C0 1px inset;BORDER-TOP: #C0C0C0 1px inset; BORDER-BOTTOM: #C0C0C0 1px inset; BACKGROUND-COLOR: transparent;"
						End If
						
					    if Trim(dailyTotalsArray(i)) <> "" then
						%>
							<td  nowrap align=right style="<%= myStyle %>"><font size="1"><%= FormatNum(Trim(dailyTotalsArray(i))) %></font></td>		
						<%
						end if
			next
				
				%>
		</tr>
        
     <%else%>
                <script language = 'javascript'>
                		alert ("No contracts found using the specified criteria");
                		window.parent.history.go(-1);          		
                </script>
                
                <%  Set Rs = Nothing
					Set Conn = Nothing
                Response.end
     
   End if
 %>

  </table>

</body>

</html>