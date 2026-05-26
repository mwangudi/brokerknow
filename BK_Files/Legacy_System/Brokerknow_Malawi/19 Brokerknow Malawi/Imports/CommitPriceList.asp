	
	<!--#include virtual="libroutines.asp"-->
	<%

		Set Rs = Server.CreateObject("ADODB.Recordset")			
		Set Rs1 = Server.CreateObject("ADODB.Recordset")			
		Set Conn = Server.CreateObject("ADODB.Connection")
		Set Conn = GetActiveConnection("KBroker")

	'' now save the data here
		'first select the records from the _imported pricelist
		'response.write trim(request("commit")) & "sdfsdf"
	theDate = Request("theDate")
	
	if trim(request("commit"))<>"" then

		sql = "SELECT _ImportPriceList.SecKnow_DPA_, isnull(_ImportPriceList.price,0) as price, _ImportPriceList.importdate, Security.SecurityCode FROM _ImportPriceList INNER JOIN Security ON _ImportPriceList.SecKnow_DPA_ = Security.Security_DPA_ order by secknow_DPA_"

		rs.open sql, conn, 0,1
		
		conn.BeginTrans
		'first delete only the Price list imported for that day if any b4  adding the new records
		'sqlstrDel = "Delete from Datastream_Market where MktDate = '" & formatdate(now())& "'"
		sqlstrDel = "Delete from Datastream_Market where MktDate = '" & formatdate(theDate)& "'"
		sqlstrDel = SQLServerFormat(HandleQuote(sqlstrDel))
		conn.execute (sqlstrDel)
		while not rs.eof 
		'Add record to Datastream Market
		 if rs("price")<>0 then
			sqlStr = "INSERT INTO [Datastream_Market] (MktUnique, MktClose, MktCode, MktDate) " & _
				  " SELECT " & " " & "iif(isnull(max([MktUnique])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Datastream_Market'),max([MktUnique]) + 1)" & " " & " as MktUnique" & _
				 "," & " " & rs("price") & " " & " as MktClose  " & _
				  "," & "'" & rs("SecurityCode") & "'" & " as MktCode" & _
				  "," & "#" & FormatDate(theDate) & "#" & " as MktDate" & _
				  " FROM [Datastream_Market]"
		else
			sqlStr = "INSERT INTO [Datastream_Market] (MktUnique, MktClose, MktCode, MktDate) " & _
				  " SELECT " & " " & "iif(isnull(max([MktUnique])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'Datastream_Market'),max([MktUnique]) + 1)" & " " & " as MktUnique" & _
				 "," & " null " & " as MktClose  " & _
				  "," & "'" & rs("SecurityCode") & "'" & " as MktCode" & _
				  "," & "#" & FormatDate(theDate) & "#" & " as MktDate" & _
				  " FROM [Datastream_Market]"
		end if

			sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
			'response.write  sqlstr
			'response.end
			conn.Execute (sqlStr)
		rs.movenext
		wend
			
		conn.CommitTrans

		rs.close
		set rs=nothing
		conn.close
		set conn = nothing
		
		''BONDS AND EQUITY
		Bond = trim(replace(Request("hidBond"),",",""))
		Equity = trim(replace(Request("hidEquity"),",",""))
	
		'Response.Write Bond & "<br>"
		'Response.Write Equity & "<br>"
		'Response.End 
		
		'sBond = Split(Bond,",")
		'sEquity = Split(Equity,",")
	
		Set Conn = Server.CreateObject("ADODB.Connection")
		Set Conn = GetActiveConnection("KBroker")
	
		'SQL = "INSERT INTO _ImportBondsEquitiesTurnover (Bonds, Equities)" & _
		'"VALUES ('"& sBond(1) &"', '"& sEquity(0) &"')"
		
		'SQL = "INSERT INTO _ImportBondsEquitiesTurnover (Bonds, Equities, ImportDate)" & _
		'"VALUES ("& Bond &", '"& Equity &"', #"& FormatDate(theDate) &"#)"
		
		'Response.Write SQL
		'Response.End 
		
		'Conn.Execute SQL
		%>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			window.alert("Pricelist succesfully commited");
			window.alert("Bonds and Equity Turnovers succesfully commited.");
			window.location='import.asp';
		//-->
		</SCRIPT>
		<%
		end if
	%>