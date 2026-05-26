<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Un Matched Holdings</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	<!--CALENDAR -->
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>

	<style media="print">	
		@page domes {
			size: portrait;
			margin-left: 2cm;
			margin-right: 5cm;
			margin-top: 1cm;    
			margin-bottom: 2cm;
			br.newpage{
				page-break-before:always;
			}
		}
	</style>


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
	window.onload = HideRemindSelectLandscape;
	//window.onbeforeprint = HideRemindSelectLandscape;
	//window.onafterprint = ShowRemindSelectLandscape;
</Script>

<!--#include file="../libroutines.asp"-->

<%
	Dim conn 
	Dim sqlStr
	Dim rs
	
    sqlStr = "select CDSNo, SecurityImportcode, Quantity, Accountstatus from CDSUnMatchedHoldings"
		 
    Set conn = GetActiveConnection("KBroker")
		
    Set Rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    
    intrscount = rs.recordcount
    
    if intrscount <= 0 then
     %>
		<Script Language="JavaScript">
			alert("No unmatched holdings available.");					
        </Script>
     <% Set Rs = Nothing
        Set Conn = Nothing
        Response.End
    end if
    
		        
Rs.MoveFirst
		  
rsdata = rs.getrows() 'get data into an array
	        
selectedTradeDate = FormatDate(Date)

DrawPageFunctions True, True, False

headerDescription = FormatDateFull(selectedTradeDate) %>
<i id="landRem"></i>
<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4">Un Matched Holdings</font></b></td>
		<td nowrap align=right><font face="Impact" size="3"><%= Session("CompanyName") %></font></td>
	</tr>	   
	<tr>
		<td nowrap colspan="2">as at <%=selectedTradeDate%></td>
	</tr> 
    <tr>
		  <td COLSPAN=2><font face="Arial" size="2">&nbsp;</font></td>
	</tr>
</table>				

  <table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="600">
    <tr>     
     <td bgcolor="#000000" ><b><font color="#FFFFFF" face="Arial Narrow">CDS Number</font></b></td>
     <td bgcolor="#000000" ><b><font color="#FFFFFF" face="Arial Narrow">Security</font></b></td>
     <td bgcolor="#000000" ><b><font color="#FFFFFF" face="Arial Narrow">Quantity</font></b></td>
     <td bgcolor="#000000" ><b><font color="#FFFFFF" face="Arial Narrow">Account status</font></b></td>         		
	</tr>
    <tr>
      <td colspan="4" width="600">&nbsp; </td>
    </tr>        
    <%    
    
    for intcount = 0 to intrscount-1
		%>
			<tr>		  
			  <td><font face="Arial Narrow"><%=trim(rsdata(0,intcount))%></font></td>      
			  <td><font face="Arial Narrow"><%=trim(rsdata(1,intcount))%></font></td>
			  <td><font face="Arial Narrow"><%=FormatNumCommasOnly(trim(rsdata(2,intcount)))%></font></td>
			  <td><font face="Arial Narrow"><%=trim(rsdata(3,intcount))%></font></td>   
			</tr>		
		 <%	
    next

    Set Rs = Nothing
    Set Conn = Nothing
    %>

    <tr>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="4">&nbsp;</td>
    </tr>
  </table>
  

</body>

</html>