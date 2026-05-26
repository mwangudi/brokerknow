<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Trading Schedule</title>
 
	 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
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
</head>

<body Class="Reports">


<!--#include file="../libroutines.asp"-->

<%
	Dim conn 
   Dim sqlStr
   Dim rs
	
		 sqlStr = "SELECT * FROM [TradingSchedule]"
		 
		 Set conn = GetActiveConnection("KBroker")
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then %>
				<Script Language="JavaScript">
					alert("No orders available");
					window.parent.close();					
                </Script>
                <% Set Rs = Nothing
                Set Conn = Nothing
                Response.End
        End If
        
        rs.MoveFirst
        
%>

<% DrawPageFunctions True, True, False %>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		     <td>
		        <b><font face="Arial Narrow" size="4">Trading Schedule</font></b></td>
		      <td align=right>
				<b><font face="Arial Narrow" size="4"><%= Session("CompanyName") %></font></b></td>
					
			</td>  
		</tr>	
     <tr>
		     <td colspan=2>
		        <font face="Arial" size="2">for Deals traded on: <%= FormatDateFull(Date) %></font></td>
		</tr>
    <tr>
    
     <tr>
		     <td colspan=2>
		        <font face="Arial" size="2">&nbsp;</font></td>
		</tr>
</table>    

<table border="0" cellspacing="0" cellpadding="4" style="font-family: Arial Narrow; LEFT-MARGIN:100PX">

<tr>
	<td nowrap bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Order No</font></b></td>
	<td nowrap bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Type</font></b></td>
	<td nowrap bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Security</font></b></td>
	<td nowrap bgcolor="#000000"><b><font color="#FFFFFF" face="Arial" size="2">Date of Order</font></b></td>		
	<td nowrap bgcolor="#000000" align="right"><b><font color="#FFFFFF" face="Arial" size="2">Quantity</font></b></td>
	<td nowrap bgcolor="#000000" align="right"><b><font color="#FFFFFF" face="Arial" size="2">Price</font></b></td>
	
  </tr>
<%		Do Until rs.EOF%>
                <tr>
                        <td nowrap><%=rs.Fields("Order_DPA_")%></td>
                        <td nowrap><%=rs.Fields("OrdDetailType")%></td>
                        <td nowrap><%=rs.Fields("OrdDetailSecurity")%></td>
                        <td nowrap><%=FormatDate(rs.Fields("OrderDate"))%></td>                        
                        <td nowrap align="right"><%=FormatNumCommasOnly(rs.Fields("BalanceQty"))%></td>
                        <td nowrap align="right"><%=FormatNum(rs.Fields("OrdDetailPrice"))%></td>
                        
                </tr>
                <%rs.MoveNext
        Loop
        conn.Close
        Set conn = Nothing%>
</table>


</body>

</html>
