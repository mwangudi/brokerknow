<%
cID = Request.QueryString("cID")
sID = Request.QueryString("sID")

If cID <> "" Then
	Set conn = CreateObject("ADODB.Connection")

	'Response.Write "FILE NAME=" & server.MapPath("../") & "\UDL\KBroker.UDL"
	'Response.End 
	
    conn.ConnectionString = "FILE NAME=" & server.MapPath("../") & "\UDL\KBroker.UDL"
    conn.Open
    
	if sID = "" then sID = 0
	
	SqlStr = "SELECT IsNull(Quantity,0) AS Quantity FROM IPOHoldings " & _
	" WHERE (Security_DPA_ = "& sID &") AND (Client_DPA_ = "& cID &")"
	Set Rs = conn.Execute(SqlStr)
	
	If Not (Rs.EOF Or Rs.BOF) Then
		str = Rs("Quantity")
	Else
		str = "0"
	End If
	
	Response.Write str
End If
%>