<!--#include file="../libroutines.asp"-->
<%
	''UPDATE COMMISSIONS FOR COMPOUNDED CONTRACTS	
    
	Set conn = GetActiveConnection("KBroker")
		 
	''RUN STORED PROCEDURE
	Conn.execute("UpdateCompoundedContractCommissions")	
	        
%>

<SCRIPT LANGUAGE="JavaScript">
	ShowMessage('Compounded contracts updated');
	window.location.replace("import.asp");
</SCRIPT>