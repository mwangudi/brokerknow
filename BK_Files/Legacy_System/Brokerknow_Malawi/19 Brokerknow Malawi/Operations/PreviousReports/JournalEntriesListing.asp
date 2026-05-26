<html>

<head>
<meta http-equiv="Content-Language" content="en-uk">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Journal Entries Listing</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

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

<!--#include file="../libroutines.asp"-->

<%


genReport = Request.Form("genReport")

If genReport = "" Then%>
	<Script Language="JavaScript">
		document.body.className = 'dialog';
		
		function validateForm(frm){						
			frm.target = '_self';			
			frm.submit();
		}	
	
	</Script>
	<form method="POST" action="JournalEntriesListing.asp" Name="frmMain" id="frmMain">		
		<input type="hidden" name="genReport" value="1">
			<tr>
				<td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.getElementById('frmMain'))" Value=" Generate... ">&nbsp;&nbsp; <input type="Button" class="Buttons" Value=" Close " OnClick="JavaScript: window.parent.self.close();"></td>
			</tr>
		</table>
		
	</form>
	
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>


<% DrawPageFunctions True, True, True

sqlStr = "JournaEntriesProc"




Set Conn = GetActiveConnection("KBroker")

Set Rs = Conn.Execute(sqlStr)

 If Rs.EOF Or Rs.BOF Then%>
		<Script Language="JavaScript">	
			ShowMessage('No data was found using the criteria entered.');
			window.parent.history.go(-1);
		</Script>
		<%Set Conn = Nothing
		Set Rs = Nothing
		Response.End
  End If


 %>

<table border="0" cellspacing="0" cellpadding="2" style="font-family: Arial Narrow" width="100%">
     <tr>
		<td nowrap><b><font face="Arial Narrow" size="4"><%= report_description %></font></b></td>
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
	
	<% Do Until rs.EOF%>
        		<tr>
      				
      				<%For i = 0 To Rs.Fields.Count - 1%>		
					 	 <td nowrap><%= Rs.Fields(i).Value %></td>
				   <%Next %>			
      
	        </tr>
	 <%  Rs.MoveNext
	 Loop%>	  	
  </table>
  
	 <%	 
	 Set Rs = Nothing
	 Set Conn = Nothing
     %>	
</body>

</html>
