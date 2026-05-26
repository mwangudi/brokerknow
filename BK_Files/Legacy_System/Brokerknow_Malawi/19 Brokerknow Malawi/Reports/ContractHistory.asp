<html>
<%
Set conn = GetActiveConnection("KBroker")
%>
<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Contract History</title>
	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
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
			
		}

	</style>
	<script language="javascript">
	 function  toggleClient(chkName){
	  
	  if (chkName.checked==true){
	    document.getElementById("txtClientCode").style.display = ""
	    document.getElementById("cboClient").style.display = ""
	   }else{
	    document.getElementById("txtClientCode").style.display = "none"
	    document.getElementById("cboClient").style.display = "none"
	   }
	 
	 }
	</script>
</head>

<body Class="Reports">
<!--#include file="../libroutines.asp"-->

<%
FirstDay=DateSerial(Year(Date), Month(Date)-1 + iOffset, 1)

genReport = trim(Request.Form("genReport"))
ContractType = trim(Request.Form("cboContractType"))
selectedFromDate = trim(Request.Form("transFromDate"))
selectedToDate = trim(Request.Form("transToDate"))
filterClient = trim(Request.Form("chkclient"))
Client = trim(Request.Form("cboClient"))

If genReport <> "1" Or selectedFromDate = ""  Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
		function validateForm(frm){	
		 		
			if (frm.transFromDate.value ==''){
				alert("Please specify the start date.");
				return;
			}
			
			if (frm.chkclient.checked==true){
			 if (frm.cboClient.value=='') {
			     alert("Please specify the client.");
				 return;
			 }
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		
		var cal=new ctlSpiffyCalendarBox("cal", "frmMain", "transFromDate","cmdDate","<%= FormatDate(FirstDay) %>",1);
		var cal1=new ctlSpiffyCalendarBox("cal1", "frmMain", "transToDate","cmdDate","<%= FormatDate(Date()) %>",1);
		
	</Script>
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="ContractHistory.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">
		<table>
			<tr>
				<td>Select Contract Type </td>
				<td>
				    <select name = 'cboContractType' id = 'cboContractType' size="1">					
                       <option value="0">All</option>
                       <option value="1">Deleted</option>
					</select>
				</td>
			</tr>
			<tr>
				<td>Select date from:</td>
				<td>
					<SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td>Select date To:</td>
				<td>
					<SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT>	
				</td>
			</tr>
			<tr>
				<td width="20%">
				&nbsp;<input value='1'  type="checkbox" name="chkclient" id="chkclient" onclick="toggleClient(this)"><label for="chkclient">&nbsp;Filter by Client</label>
				</td>
				<td width="80%" nowrap valign="bottom"><input style="display: none;" type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);"> 
				&nbsp;&nbsp;&nbsp;
				<select style="display: none;" name = "cboClient" id = "cboClient" size="1" 
    				onchange = "UpdateCode(true,cboClient,txtClientCode)"
					onKeypress = "return (dodefaultaction()==''); "  
					onKeydown = "return (dodefaultaction()==''); " 
					onKeyup = "change(cboClient,0);"
					onKeyup = "return (change(cboClient,0));"
				  onfocus = "txtval = '';inputIsItemCode = 0;" 
					onblur = "txtval = '';inputIsItemCode = 0;">
					<option selected SearchCode = "" SearchText = ""  value = ""></option>
					<%
					        
					        sqlStr = "SELECT Client_DPA_,ClientName FROM Client WHERE Deleted <> 1 ORDER BY ClientName"
					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        
					        intrscount = rs.recordcount
					        
					        if intrscount > 0 then
					         rs.movefirst
					         rsdata= rs.getrows()
					         
					         for intcount = 0 to intrscount-1
					          ClientID = trim(rsdata(0,intcount))
					          ClientName = trim(rsdata(1,intcount))
					          %>
					             <option SearchCode = "<%=ClientID%>" SearchText = "<%=ClientName%>"  value = '<%=ClientID%>'><%=mid(ClientName,1,50)%></option>
					          <%
					         next
					         
					        end if    
					%>

					    </select>
				</td>
			</tr>
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">&nbsp;&nbsp;</td>
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
   
   ContractType = cint(ContractType)   
   ContractSql = ""
   
   if ContractType = 1 then
    ContractSql = " Where dbo.Lot.deleted = 1 "
   end if
   
   if selectedFromDate <> "" AND  selectedToDate <> "" then
    DateSql = " (cast(floor(cast(dbo.Lot.LotTDate as numeric)) as datetime) between #" & formatdate(selectedFromDate) & _
              "# AND #" & formatdate(selectedToDate) & "#)  "
          
   else
    DateSql = " (cast(floor(cast(dbo.Lot.LotTDate as numeric)) as datetime) = #" & formatdate(selectedFromDate) & "#  " 
   end if
   
   if ContractSql = "" then DateSql = " Where " & DateSql else DateSql = " AND " & DateSql
   
   if filterClient = "" then filterClient = 0
   
   if filterClient = 1 then
      ClientSQL = " AND dbo.Client.Client_DPA_ = " & Client
   end if
   
   sqlStr = "SELECT     dbo.tbOrder.Order_DPA_, dbo.tbOrder.OrderDate, dbo.Security.SecurityCode, dbo.Security.SecurityName, dbo.Client.Client_DPA_,  " & _
			"                       dbo.Client.ClientName, dbo.Client.ClientCDSNo, dbo.OrdDetail.OrdDetailPrice, dbo.OrdDetail.OrdDetailQty, dbo.Lot.ContractNumber, dbo.Lot.LotPrice,  " & _
			"                       dbo.Lot.LotQty, dbo.Lot.LotSlipNo, dbo.Lot.LotTDate, isnull(dbo.Lot.Deleted,0) as Deleted, dbo.Lot.TimeChanged AS LastModified,  " & _
			"                       dbo.Users.Surname + ' ' + dbo.Users.OtherNames AS ModifiedBy,  dbo.OrdDetail.Best " & _
			" FROM         dbo.Lot INNER JOIN " & _
			"                       dbo.OrdDetail ON dbo.Lot.OrdDetail_DPA_ = dbo.OrdDetail.OrdDetail_DPA_ INNER JOIN " & _
			"                       dbo.tbOrder ON dbo.OrdDetail.Order_DPA_ = dbo.tbOrder.Order_DPA_ INNER JOIN " & _
			"                       dbo.Client ON dbo.tbOrder.Client_DPA_ = dbo.Client.Client_DPA_ INNER JOIN " & _
			"                       dbo.OrderType ON dbo.tbOrder.OrderType_DPA_ = dbo.OrderType.OrderType_DPA_ INNER JOIN " & _
			"                       dbo.Security ON dbo.OrdDetail.Security_DPA_ = dbo.Security.Security_DPA_ LEFT OUTER JOIN " & _
			"                       dbo.Users ON dbo.Lot.ChangedBy = dbo.Users.UserID " & ContractSql & DateSql & ClientSQL & _
			" ORDER BY dbo.Client.ClientName "
	
	set Rs = conn.execute(SQLServerFormat(HandleQuote(sqlStr)))

	intrscount = rs.recordcount
	
	if intrscount <= 0 then
	 %>
		<Script Language="JavaScript">
			alert("No records were found using the specified criterion.")
			window.location.href='ContractHistory.asp';
		</Script>
	 <%
	 Set Rs = Nothing
	 Set Conn = Nothing
	 Response.End
	end if
	
	if selectedToDate = "" then selectedToDate= Date()
	
	DateTitle = " between " & formatdate(selectedFromDate) & " AND " & formatdate(selectedToDate)
	
	
%>	
<table border="0" cellspacing="0" cellpadding="0" style="font-family: Arial Narrow" width="1070">
	<br>
	<tr>
		<td width="535" align="left"><font face="Impact" size="4">CONTRACT HISTORY</font></td>  
		<td width="535"  align="right"><font face="Impact" size="4"><%=ucase(session("CompanyName"))%></font></td>       
	</tr>
	<tr>
		<td   height="0" align="right">&nbsp;</td>      
	</tr>
	<tr>
		<td   height="0" align="left">&nbsp;<b>for the period <%=DateTitle%></b></td>      
	</tr>
	<tr>
		<td   height="0" align="right">&nbsp;</td>      
	</tr>
</table>
<br>  
<table border="0" cellspacing="0" cellpadding="0" style="font-family: Arial Narrow" width="1070">
	<tr>
		<td width="50" align="left"><b>Order No</b></td> 
		<td width="70" align="left"><b>&nbsp;Order Date</b></td>
		<td width="50" align="right"><b>O Price</b>&nbsp;</td> 
		<td width="50" align="right"><b>O Qty</b>&nbsp;</td>  
		<td width="70" align="left"><b>Traded</b></td> 
		<td width="50" align="left"><b>Contract</b></td> 
		<td width="50" align="left"><b>Slip No</b></td> 
		<td width="30" align="left"><b>Type</b></td> 
		<td width="300" align="left"><b>Security</b></td> 
		<td width="50" align="right"><b>Qty</b>&nbsp;</td> 
		<td width="50" align="right"><b>Price</b>&nbsp;</td> 
		<td width="110" align="left"><b>Last Modified</b></td> 
		<td width="100" align="left"><b>Modified By</b></td> 
		<td width="40" align="center"><b>Deleted</b></td>  
	</tr>
	
<%

dim currentClient,PeviousClient
currentClient = ""
PeviousClient =  ""

rs.movefirst
rsdata = rs.getrows()

for intcount = 0 to intrscount -1 

	currentClient = trim(rsdata(5,intcount))
	ClientID = trim(rsdata(4,intcount))
	CDSNo = trim(rsdata(6,intcount))
	OrderNo = trim(rsdata(0,intcount))
	OrderDate = formatdate(trim(rsdata(1,intcount)))
	OrdPrice = trim(rsdata(7,intcount))

	if isnumeric(OrdPrice) then
	OrdPrice = formatnum(cdbl(OrdPrice))
	end if

	OrdQty = FormatNumCommasOnly(Cdbl(trim(rsdata(8,intcount))))
	TradeDate = formatdate(trim(rsdata(13,intcount)))
	ContractNo = trim(rsdata(9,intcount))
	SlipNo = trim(rsdata(12,intcount))
	OrdType = left(ContractNo,1)
	Security = trim(rsdata(3,intcount))
	Qty = FormatNumCommasOnly(trim(rsdata(11,intcount)))
	Price = formatnum(trim(rsdata(10,intcount)))
	LastModified = trim(rsdata(15,intcount))
	Modified = trim(rsdata(16,intcount))
	Deleted = trim(rsdata(14,intcount))
    
	if Deleted  then Deleted = "Yes" else Deleted = "No"

	'row title: group data by client

	if currentClient  <> PeviousClient then
	 %>
	  <tr>
			<td  colspan="14">&nbsp;</td> 
	  </tr>
	  <tr>
			<td align="left" colspan="14"><b>&nbsp;<%=ClientID%>&nbsp;&nbsp;<%=CDSNo%>&nbsp;&nbsp;<%=currentClient%></b></td> 
		</tr>
	  <tr>
			<td colspan="14">&nbsp;</td> 
		</tr>
	 <%
	end if
	
	%>
	 <tr>
		<td align="left"><%=OrderNo%></td> 
		<td align="left"><%=OrderDate%></td>
		<td align="right"><%=OrdPrice%>&nbsp;</td> 
		<td align="right"><%=OrdQty%>&nbsp;</td>  
		<td align="left"><%=TradeDate%></td> 
		<td align="left"><%=ContractNo%></td> 
		<td align="left"><%=SlipNo%></td> 
		<td align="center"><%=OrdType%></td> 
		<td align="left"><%=Security%></td> 
		<td align="right"><%=Qty%>&nbsp;</td> 
		<td align="right"><%=Price%>&nbsp;</td> 
		<td align="left"><%=LastModified%></td> 
		<td align="left"><%=Modified%></td> 
		<td align="center"><%=Deleted%></td>  
	</tr>
	<%

PeviousClient = currentClient
next

Set Rs = Nothing
Set Conn = Nothing%>   
<tr>
		<td colspan="14">&nbsp;</td> 
	</tr>
	
</table>
</body>
</html>