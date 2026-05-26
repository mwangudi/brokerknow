<% Response.expires=-1%>
<html>
<head>
<title>Batched Applications</title>
<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css">
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

<style media="print">
	@page {
		@top{
			font-family: Helvetica, Arial, sans-serif;
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


	}

</style>
</head>

<body Class="Reports">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<!--#include file="../libroutines.asp"-->

<%
FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)
selectedType=request("cboType")
If len(Request.QueryString) > 0 Then
	genReport = 1
	IDs = Trim(Request.QueryString("From"))
	ID = Split(IDs,"<->")
	selectedFrom = ID(0)
	selectedTo = SelectedFrom
	offering = ID(1)
Else
	genReport = Request.Form("genReport")
	selectedFromDate = Request("txtDate")
	selectedToDate = Request("txtToDate")
	offering = Trim(Request.Form("cboOfferings"))

	If genReport <> "1"  Then%>
		<Script Language="JavaScript">
			report_SetBodyClass();
			function validateForm(frm){
				if (frm.cboOfferings[frm.cboOfferings.selectedIndex].SearchCode == ''){
					alert("Please select an Offering.");
					frm.cboOfferings.focus();
					return;
				}

				/*if (frm.txtFrom.value == ""){
					alert("Please specify a From Batch Number.");
					frm.txtFrom.focus();
					return;
				}

				if (frm.txtTo.value == ""){
					alert("Please specify a To Batch Number.");
					frm.txtTo.focus();
					return;
				}

				if (parseInt(frm.txtFrom.value) > parseInt(frm.txtTo.value)){
					alert("Invalid batch selection criteria.");
					frm.txtFrom.focus();
					return;
				}*/

				frm.target = '_self';
				frm.submit();

			}
			var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
			var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "txtToDate","cmdToDate","<%= FormatDate(Date) %>",1);


		</Script>
		<form method="POST" action="BatchTotals.asp" Name="frmMain" id="frmMain">
			<input type="hidden" value="1" name="genReport">
			<table>
				<tr>
					<td nowrap>Select an Offering&nbsp;</td>
					<td>
					<select name = 'cboOfferings' id = 'cboOfferings' size="1">
						<option selected SearchCode = '' value = ''></option>
						<%
						Set conn = GetActiveConnection("KBroker")
						sqlStr = "SELECT * FROM [SecurityListOfferings] WHERE Security_DPA_ = 157" & _
						" Order By DefaultSelection DESC, Security_DPA_ DESC"
						
						sqlStr = "SELECT DISTINCT  " & _
                                 "                       SecurityName, SecurityMktPrice, SecurityCode, SecurityAddr, Security_DPA_, BatchSize, OfferType, ParentSecurity_DPA_, DefaultSelection,  " & _
                                 "                       MinimumQty, Ratio, StepQty " & _
                                 " FROM         SecurityListOfferings " & _
                                 " ORDER BY DefaultSelection DESC, Security_DPA_ DESC"
						
						sqlStr = "SELECT SecurityCode,Security_DPA_,SecurityName,isnull(DefaultSelection,0) as DefaultSelection FROM [SecurityListOfferings] " & _
								 " Order By SecurityName ASC"
						
						Set rs = conn.Execute(SqlStr)

						If Not (rs.EOF Or rs.BOF) Then
							Do Until rs.EOF
								if(Cint(rs("DefaultSelection"))=1) then
								%>
								<option selected SearchCode = '<%=rs("SecurityCode")%>' value = '<%=rs("Security_DPA_")%>'><%=rs("SecurityName")%></option>
								<%
								else							
								%>
								<option SearchCode = '<%=rs("SecurityCode")%>' value = '<%=rs("Security_DPA_")%>'><%=rs("SecurityName")%></option>
								<%
								end if
							rs.MoveNext
							Loop
						End If%>
					</select>
					</td>
				</tr>
			
				<!--<tr>
					<td nowrap>From Batch No:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td><input type='text' id='txtFrom' name='txtFrom' size='10'></td>
					<td>&nbsp;</td>
					<td nowrap>To Batch No: </td>
					<td><input type='text' id='txtTo' name='txtTo' size='10'></td>
				</tr>-->
				<tr>
					<td>Application Type</td>
					<td > 
					<select name = "cboType">
					<option selected value="1"> Application</option>
					<option  value="2">Forwards</option>
					</select>
					</td>

				</tr>
			<tr>
				<td width="10%" nowrap>Date From:</td>
				<td>
				<SCRIPT LANGUAGE="JavaScript">cal.writeControl();	</SCRIPT>
				</td>
				
			</tr>
			<tr>
				<td width="10%" nowrap>Date To :</td>
				<td>
				<SCRIPT LANGUAGE="JavaScript">cal1.writeControl();	</SCRIPT>
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
End If

If len(Request.QueryString) > 0 Then
	DrawPageFunctions True, False, False
Else
	DrawPageFunctions True, True, True
End If
selectedBrokerCode = Trim (Request.Form("brokerCodes"))
'response.write selectedBrokerCode : response.end
If selectedBrokerCode ="" Then selectedBrokerCode="0"

'response.write selectedBrokerCode : response.end

Set conn = GetActiveConnection("KBroker")
Set Rs = CreateObject("ADODB.Recordset")
if trim(selectedType) = 1 or trim(selectedType) = "1" then
		

		  sqlStr = "Select  Batch_No, SecurityName, Count(Batch_No) as Num, SUM(Alloted_Rights) AS Shares, SUM(Alloted_Rights) * Offering_Price AS Total " & _
			 " FROM         OfferingList " & _
			 " WHERE     (Offering = " & offering & ") AND (Deleted <> 1) AND (CAST(FLOOR(CAST(Offerings_Date AS float)) AS datetime) BETWEEN '" & selectedFromDate & "' AND '" & selectedToDate & "') " & _
			 " GROUP BY Offering_Price, Batch_No, SecurityName"



 else
		sqlStr= "SELECT     Batch_No, SUM(Alloted_Rights) AS Shares, SUM(Extra) AS Additional, Offering_Price, SUM(Alloted_Rights) + SUM(Extra) AS TotalShares,  " & _
		 "                       (SUM(Alloted_Rights) + SUM(Extra)) * Offering_Price AS Total, SecurityName " & _
		 " FROM         ForwardsList " & _
		 " WHERE     (Batch_No BETWEEN " & selectedFrom & " AND " & selectedTo  & ") AND (Offering = " & offering & ") " & _
		 " GROUP BY Offering_Price, Batch_No, SecurityName"

		
		sqlStr= "SELECT    		  Batch_No, SUM(Alloted_Rights) AS Shares, Offering_Price, SUM(Alloted_Rights) AS TotalShares, SUM(Alloted_Rights) * Offering_Price AS Total, SecurityName,  " & _
		 "                       (SUM(Alloted_Rights) + SUM(Extra)) * Offering_Price AS Total, SecurityName " & _
		 " FROM         ForwardsList " & _
		 " WHERE     (Batch_No BETWEEN " & selectedFrom & " AND " & selectedTo  & ") AND (Offering = " & offering & ") " & _
		 " GROUP BY Offering_Price, Batch_No, SecurityName"


		  sqlStr = " Select  Batch_No, SecurityName, Count(Batch_No) as Num, SUM(Alloted_Rights) AS Shares, SUM(Alloted_Rights) * Offering_Price AS Total " & _
			 " FROM         ForwardsList " & _
			 " WHERE     (Offering = " & offering & ") AND (Deleted <> 1) AND (CAST(FLOOR(CAST(Offerings_Date AS float)) AS datetime) BETWEEN '" & selectedFromDate & "' AND '" & selectedToDate & "') " & _
			 " GROUP BY Offering_Price, Batch_No, SecurityName"
		
 end if
'response.write sqlStr : response.end
Rs.CursorLocation = adUseClient
Rs.Open (sqlStr), conn.ConnectionString, adOpenKeyset, adLockOptimistic

If rs.EOF Or rs.BOF Then
	%>
	<Script Language="JavaScript">
		alert("The report did not find any values available.")
		window.history.go(-1);
	</Script>
	<%Set Rs = Nothing
	Set Conn = Nothing
	Response.End
Else
	offeringName = Trim(Rs("SecurityName"))
	
End If

	'output single batch
	%>
	<table border="0" cellspacing="2" cellpadding="3" style="font-family: Arial Narrow" width="100%">
		<tr>
		  <table border ="0" width ="100%">
		   <tr>
			<td><font face="Impact" size="3"nowrap>BATCH SUMMARY</font></td>
			<td align="right" width="70%" nowrap align=right><font face="Impact" size="3"><%= UCase(Session("CompanyName")) %></font></td>
			</tr>
			<tr><td align="center" width="100%" nowrap colspan='2'><font face="Impact" size="5"><%=UCase(offeringName)%></font></td></tr>
			<tr><td align="center" width="100%"colspan='2'><%= UCase(offeringName) & "&nbsp;BATCH CONTROL SCHEDULE"  %></font></td></tr>
		  </table>
		</tr>
		<!--tr>
			<td align="center" width="100%" nowrap><font face="Impact" size="5"><%'=UCase(offeringName)%></font></td>
		</tr>
		<tr>
		  <td align="center" width="100%"><%'= UCase(offeringName) & "&nbsp;BATCH SUMMARY" %></font></td>
		</tr-->
	</table>
	<table border="0" cellspacing="0" cellpadding="3"  style="font-family: Arial Narrow" width="100%">
		 <tr>
		  	  
		   <td style="border-left: 1px solid black;border-top: 1px solid black;border-right: 1px solid black;border-top: 1px solid black" width="10%"><B>BATCH&nbsp;NO</B></td>	
		     <td style="border-right: 1px solid black;border-top: 1px solid black" ><B>N0. OF APPLICATIONS</B></td>
		   <td style="border-right: 1px solid black;border-top: 1px solid black" width="15%"><B>TOTAL NO.&nbsp;OF&nbsp;SHARES</B></td>
   		   <!-- <td style="border-right: 1px solid black;border-top: 1px solid black" width="15%"><B>NO.&nbsp;OF&nbsp;ADDITIONAL&nbsp;SHARES</B></td>	 -->	   
		 
		   <td style="border-right: 1px solid black;border-top: 1px solid black" width="15%"><B>AMOUNT&nbsp;PAYABLE</B></td>
		</tr>
		<%
		totalquantity=0
		totalpayable=0
		totalApps=0
		num=1
		do while rs.eof=false
		%>
		<tr>		  	  
		   <td style="border-left: 1px solid black;border-top: 1px solid black;border-right: 1px solid black;border-top: 1px solid black" width="10%"><%=rs("Batch_No")%></td>		   
		   <td  style="border-right: 1px solid black;border-top: 1px solid black" width="10%"><%=rs("NUM")%></td>
   		    
		   <td align ="right"  style="border-right: 1px solid black;border-top: 1px solid black" width="15%"><%=formatnumber(rs("Shares"),0)%></td>
		   <td align ="right" style="border-right: 1px solid black;border-top: 1px solid black" width="15%"><%=formatnum(rs("Total"))%></td>
		</tr>
		<%
		totalApps = totalApps + rs("NUM")
		totalquantity=totalquantity + rs("Shares")
		totalpayable=totalpayable + ccur(rs("Total"))
		rs.movenext
		loop
		%>
		<tr>
			
			
			<td   ALIGN="RIGHT"style="border-left: 1px solid black;border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" ><b>TOTALS</b></td>
			<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" align="left" width="15%"><b><%=FormatNumEx(totalApps,0)%></b>&nbsp;</td>
			<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" align="right" width="15%"><b><%=FormatNumEx(totalquantity,0)%></b>&nbsp;</td>
			<td style="border-right: 1px solid black;border-top: 1px solid black;border-bottom: 1px solid black" align="right" width="15%"><b><%=FormatNumber(totalpayable,2)%></b>&nbsp;</td>
		
		</tr>
	</table>

	
<%
Set Rs = Nothing
Set Conn = Nothing%>
</body>

</html>

