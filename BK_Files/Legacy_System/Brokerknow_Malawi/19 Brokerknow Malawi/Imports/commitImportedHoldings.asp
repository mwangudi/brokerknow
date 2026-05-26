<!--#include virtual="libroutines.asp"-->

<%	
		const UDLName = "KBroker"
		const DataSource = "CDSMatchedHoldings"
		const DataEntity = "CDSImportHolding"
		const DataEntityPlural = "CDSImportHoldings"
		const ActionFolder = "Import"
		
'======================= End_Alter_Across_Entities =================================		
		
		Dim conn 
		Dim sqlStr
		Dim Rs
		Dim Rst
		
		Set Conn = Server.CreateObject("ADODB.Connection")
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Set Rst = Server.CreateObject("ADODB.Recordset")
		Set Conn = GetActiveConnection(UDLName)
		        
		
	'delete the already commited holdings
		Conn.Execute("Delete From Holdings")
		
		sql= "select * from CDSMatchedHoldings where Imported <> 1"	
		Rst.Open sql,Conn.ConnectionString, 3, 2	
		if Rst.eof or Rst.Bof then 
		%>
		<script language="javascript">
			alert('There are currently no holdings to commit')
			window.location.replace("importHoldings.asp")
		</script>
		<%
			
		response.end
		end if
		sqlStr = "select * from Holdings "
		Rs.Open  SQLServerFormat(HandleQuote(sqlStr)), Conn.ConnectionString, 3, 2
		'add the fields to the holdings table
		Rst.MoveFirst 
		dim clientDPA
		dim securityDPA
		dim rst1
		set rst1 = server.CreateObject("Adodb.Recordset")
		while not Rst.EOF 
			rst1.Open "select Client_DPA_ from client where ClientCDSNo= '" & rst("CDSNo") & "'",conn,3,2
				rst1.MoveFirst 
				clientDPA =rst1("Client_DPA_").Value 
			rst1.Close
			'check whether the security is a right and add tha to the security code
			if rst("isRight")=true then
				secCode =rst("securityCode") & "R"
				if secCode ="UCHMR" then secCode="UCHUMI RIGHTS"
				'rs("isRight")=1
				'response.write secCode
			else
				secCode =rst("securityCode")	
				'rs("isRight")=0
			end if

			if rst("isRight")=true then
				secCode =rst("securityCode") & "R"
				if secCode ="CFCR" then secCode="CFC RIGHTS"
				'rs("isRight")=1
				'response.write secCode
			else
				secCode =rst("securityCode")	
				'rs("isRight")=0
			end if
			'response.write "Select Security_DPA_ from security where SecurityCode='"& secCode & "' <br>"
			'response.end
			rst1.Open "Select Security_DPA_ from security where SecurityCode= N'"& secCode & "'", conn,3,2
				rst1.MoveFirst 
				securityDPA =rst1("Security_DPA_").Value 
			rst1.Close 
			
			rs.AddNew 
			rs("TradeDate")=rst("TradeDate")
			rs("TradeTime")=rst("TradeTime")
			
			rs("Client_DPA_")=clientDPA
			rs("Security_DPA_")=securityDPA
			rs("Quantity")=rst("Quantity")
			rs("AccountStatus")=rst("AccountStatus")
			if rst("BalanceFree")=1 then
				rs("BalanceFree")="Y"
			else
				rs("BalanceFree")="N"
			end if
			rst("Imported")=1
			rst.update
			rs.update
		Rst.MoveNext 	
		wend	
		
		rs.Close
		Rst.Close
		set rst=nothing
		set rs=nothing
		Set Conn = Nothing
	%>
	<script language="javascript">
		alert("Holdings are succesfully commited");
		window.location.replace("importHoldings.asp")
	</script>