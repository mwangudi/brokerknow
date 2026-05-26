<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Files For Download</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	 <SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	 <SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
	

<style media="print">
	
		@page {
			margin-left: 2cm;
			margin-right: 5cm;
			margin-top: 1cm;    
			margin-bottom: 2cm;
			writing-mode: tb-rl;
			height: 80%;
			margin: 10% 0%;						
			br.newpage{
				page-break-before:always;
			}		
		}		 
		
	</style>

</head>

<body Class="Reports">

<Script Language="JavaScript">
	
	function GoBack() {
		window.location.replace("Import.asp");
	}
	
	function ShowUnmatchedTrades() {
		window.location.replace("UnmatchedTradesBody.asp");
	}
	
</Script>

<!--#include file="../libroutinesTEST.asp"-->


<table border=0 cellspacing=5 cellpadding=5 id="StatusRow" style="display:none;">
	<tr>	
		<TH noWrap align=left width="20%" bgColor=khaki><b>&nbsp;Processing next batch ...&nbsp;</b></TH>
		<td>
			&nbsp;
		</td>
	</tr>
</table>



<%

'Begin Listing Page
DrawPageFunctions True, False, False, False

 sqlStr = " SELECT     Report, Filename, CreatedByDesc AS [Creted By], right(100 + DATENAME([hour], TimeCreated),2) + ':' + right(100 + DATENAME([minute], TimeCreated),2) + ' ' + DATENAME([day], "
 sqlStr = sqlStr & " TimeCreated) + '-' + LEFT(DATENAME([month], TimeCreated), 3) AS Created " 
sqlStr = sqlStr & " FROM         down_File " 
sqlStr = sqlStr & " where (cast(floor(cast(TimeCreated as float)) as datetime) = cast(floor(cast(getdate() as float)) as datetime))" 
sqlStr = sqlStr & " Order by down_File_DPA_ Desc " 


Set Conn = GetActiveConnection("KBroker")

Set Rs = Conn.Execute(sqlStr)


 %>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap>
			<b><font face="Arial Narrow" size="4">
			Download Report
			</font></b>
		</td>
		<td nowrap>
			&nbsp;
		</td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	
       <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>			



    <table border="0" width="100%" cellPadding="2" cellSpacing=0>
    <tr bgColor="#000000">
			
	<%For i = 0 To Rs.Fields.Count - 1%>		
	 	 <td nowrap><b><font color="#FFFFFF"><%= Rs.Fields(i).Name %></font></b></td>
   <%Next %>
	</tr>
	
	<% 

	Do Until rs.EOF

			%>
				<tr>
					<%
						For i = 0 To Rs.Fields.Count - 1
							if i = 1 then
							%>
								 <td nowrap ><a href="Downloaded\<%= Rs.Fields(i).Value %>"><b><%= Rs.Fields(i).Value %><b></a></td>
							<%
							else
							%>
								 <td nowrap><%= Rs.Fields(i).Value %></td>
							<%
							end if
						Next 
					%>			
				</tr>
			<%  

	Rs.MoveNext

	Loop

	%>	  	

  </table>
  
	 <%	 
	 Set Rs = Nothing
	 Set Conn = Nothing
     %>	



<SCRIPT Language="JavaScript">
	function ShowContractSchedule()	{
		document.getElementById("cmdContractSchedule").style.display="none";
		window.location.replace("down_ContractSchedule.asp");
	}
	function ShowOutstandingOrders()	{
		document.getElementById("cmdOutstandindingOrders").style.display="none";
		window.location.replace("down_OutstandingOrderList.asp");
	}

	function ShowPurchaseOrders()	{
		document.getElementById("cmdPurchaseOrders").style.display="none";
		window.location.replace("down_PurchaseOrderList.asp");
	}
	function ShowSaleOrders()	{
		document.getElementById("cmdSaleOrders").style.display="none";
		window.location.replace("down_SaleOrderList.asp");
	}
</SCRIPT>

<table border=0 cellspacing=5 cellpadding=5>
	<tr>	
		<td>
			<INPUT type=Button  value="Generate Contract Schedule" name="cmdContractSchedule" ID="cmdContractSchedule" OnClick="JavaScript: ShowContractSchedule();">
		</td>
	</tr>
	<tr>	
		<td>
			<INPUT type=Button  value="Generate Outstanding Orders" name="cmdOutstandindingOrders" ID="cmdOutstandindingOrders" OnClick="JavaScript: ShowOutstandingOrders();">
		</td>
	</tr>
	<tr>	
		<td>
			<INPUT type=Button  value="Generate Purchase Orders" name="cmdPurchaseOrders" ID="cmdPurchaseOrders" OnClick="JavaScript: ShowPurchaseOrders();">
		</td>
	</tr>
	<tr>	
		<td>
			<INPUT type=Button  value="Generate Sale Orders" name="cmdSaleOrders" ID="cmdSaleOrders" OnClick="JavaScript: ShowSaleOrders();">
		</td>
	</tr>
</table>



</body>

</html>
