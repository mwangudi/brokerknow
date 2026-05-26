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
	<form method="POST" action="CBKBondTradeConfirmation.asp" Name="frmMain" id="frmMain">
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
   
   sqlStr = "SELECT Client.ClientAddr AS Address, CBKBondConfirmationList.*,Owner.OwnerFname + ' ' + Owner.OwnerLName AS AccountManager," & _ 
            " Client_1.ClientName AS CounterParty,Client_1.ClientOfficeTel AS CounterTel,Owner.IdNo AS IDNO FROM Client Client_1 INNER JOIN BondConfirmation ON Client_1.Client_DPA_ = BondConfirmation.CounterParty_DPA_ RIGHT OUTER JOIN " & _
            " CBKBondConfirmationList INNER JOIN Client ON CBKBondConfirmationList.Client_DPA_ = Client.Client_DPA_ INNER JOIN" & _
            " Owner ON Client.Owner_DPA_ = Owner.Owner_DPA_ ON BondConfirmation.Order_DPA_ = CBKBondConfirmationList.Order_DPA_ WHERE TransDate = '" & FormatDate(selectedTradeDate) & "'"   
  
   
   
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
				
%> 
<p class=MsoNormal style='margin-left:54.0pt'>    
<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%">
	<tr class="pageNumbering">
		<td align="left" >
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>			
    <tr>
      <td width="50%">&nbsp;</td>
      <td height = "60" valign = top ><b><%= sumRs.Fields("Type").Value %></b>&nbsp;</td>
    </tr>
    <tr>
      <td width="50%">&nbsp;</td>
      <td height = "29" valign = top ><b><%= sumRs.Fields("SlipNo").Value %>&nbsp;</b></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td height = "17" valign = top ><b><%= FormatDate(sumRs("TransDate")) %></b>&nbsp;</td>
    </tr>
    <% if(isnull(sumRs("SettlementDate"))or sumRs("SettlementDate")="") then
	%>
    <tr>
      <td>&nbsp;</td>
      <td height = "30" valign = top ><b>&nbsp;<b>&nbsp;</td>
    </tr>
	<% else
	%>
    <tr>
      <td>&nbsp;</td>
      <td height = "30" valign = top ><b><%= FormatDate(sumRs("SettlementDate")) %><b>&nbsp;</td>
    </tr>
    <% end if%>
    <tr>
      <td></td>
      <td height = "250" valign = top ><b><%=FormatNumEx(sumRs("Price"),4)%>%</b></td>
    </tr>
    <tr>
      <td width="50%">&nbsp;</td>
      <td height = "25" valign = top ><b><i>KSH.<%= FormatNum(sumRs("FaceValue")) %></i></b></td>
    </tr>
    <tr>
      <td width="50%">&nbsp;</td>
      <td height = "24" valign = top ><b><i><%=sumRs("BondIssue")%></b>&nbsp;</td>
    </tr>      
    <tr>
      <td width="50%">&nbsp;</td>
      <td height = "19" valign = top ><b><%= FormatDate(sumRs.Fields("IssueDate").Value) %></b>&nbsp;</td>
    </tr>        
    <tr>
      <td width="50%">&nbsp;</td>
      <td height = "51" valign = top ><b><%= FormatDate(sumRs.Fields("MaturityDate").Value) %></b>&nbsp;</td>
    </tr>
    <% if(sumRs("Type")="SALE") then%>      
    <tr>
      <td>&nbsp;</td>      
      <td height = "23" valign = top ><b><%= ucase(sumRs.Fields("BrokerName").Value) %></b>&nbsp;</td>
    </tr>  
    <tr>
      <td>&nbsp;</td>      
      <td height = "23" valign = top ><b><%= sumRs.Fields("BrokerCode").Value %></b>&nbsp;</td>
    </tr>    
    <tr>
      <td>&nbsp;</td>      
      <td height = "23" valign = top ><b><%= sumRs.Fields("CounterTel").Value %></b>&nbsp;</td>
    </tr>    
    <tr>
      <td>&nbsp;</td>      
      <td height = "23" valign = top ><b><%= sumRs.Fields("AccountManager").Value %></b>&nbsp;</td>
    </tr>    
    <tr>
      <td>&nbsp;</td>      
      <td height = "23" valign = top ><b><%= sumRs.Fields("IDNO").Value %></b>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>      
      <td height = "23" valign = top ><b><%= ucase(sumRs.Fields("BrokerName").Value) %></b>&nbsp;</td>
    </tr>
   <% 
   else
   %>      
    <tr>
      <td>&nbsp;</td>      
      <td height = "23" valign = top ><b><%= ucase(sumRs.Fields("BrokerName").Value) %></b>&nbsp;</td>
    </tr>  
    <tr>
      <td>&nbsp;</td>      
      <td height = "23" valign = top ><b><%= sumRs.Fields("BrokerCode").Value %></b>&nbsp;</td>
    </tr>    
    <tr>
      <td>&nbsp;</td>      
      <td height = "23" valign = top ><b><%= sumRs.Fields("BrokerOfficeTel").Value %></b>&nbsp;</td>
    </tr>    
    <tr>
      <td>&nbsp;</td>      
      <td height = "23" valign = top ><b><%= sumRs.Fields("AccountManager").Value %></b>&nbsp;</td>
    </tr>    
    <tr>
      <td>&nbsp;</td>      
      <td height = "23" valign = top ><b><%= sumRs.Fields("IDNO").Value %></b>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>      
      <td height = "23" valign = top ><b><%= ucase(sumRs.Fields("BrokerName").Value) %></b>&nbsp;</td>
    </tr>
   <% 
   end if%>
    <tr>
      <td width="50%">&nbsp;</td>
      <td>&nbsp;</b></td>
    </tr> <tr>
      <td width="50%">&nbsp;</td>
      <td>&nbsp;</b></td>
    </tr> <tr>
      <td width="50%">&nbsp;</td>
      <td>&nbsp;</b></td>
    </tr> 
    <tr>
      <td><b>Dyer and Blair Investment Bank Ltd</b></td>      
      <td><b><%= sumRs.Fields("AccountManager").Value %></b>&nbsp;</td>
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