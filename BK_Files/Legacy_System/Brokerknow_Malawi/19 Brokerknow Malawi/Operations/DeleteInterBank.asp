<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Delete Security</title>


<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

 
</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
 
<%
	const LinkedIndependent = 1
   const LinkedDependent = 2
	
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim InterBankRs
   	
	Set conn = GetActiveConnection("KBroker")
    Set InterBankRs = Server.CreateObject("ADODB.Recordset")
		
	InterBankRs.CursorLocation = adUseClient
		
        
	action = ucase(Request.Form("delAction"))
	if action = "EXECUTE" then
		  
       ID = Request.Form("ID")
				
		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for deletion"
                						   
                </script>
                <%response.end
        End If
		
		sqlStr="Select InterTransfer_DPA_ From InterTransfer Where(Contract_DPA_=" & ID & ")"
		
		Set InterBankRs=Conn.execute(sqlStr)
			if not(InterBankRs.EOF and InterBankRs.BOF) then
			ID=InterBankRs("InterTransfer_DPA_")
			else
			%>
                <script language = 'vbscript'>
                		ShowMessage "Contract has got no Inter Transfer Record"
                						   
                </script>
                <%response.end
			end if
		set InterBankRs=nothing                
                
                'delete from database
                sqlStr = "DELETE FROM [InterTransfer] WHERE InterTransfer_DPA_ = " & ID
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        'conn.CommitTrans
        Set Conn = Nothing
        WriteDeleteCloseScript
        Response.End
   	end If

		 sqlStr = "SELECT * FROM [SecurityList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then
                %><p>No Securities found</p><%
                response.end
        End If
        
        rs.MoveFirst
        
%>
<form name = 'frmDeleteSecurityList' method = 'post' action = 'DeleteSecurityList.asp' >
<table border="0" width="100%">
<tr>
	<td width="34%"><b><font color="#ff3333">Name</font></b></td>
	<td width="34%"><b><font color="#ff3333">Code</font></b></td>
	<td width="34%"><b><font color="#ff3333">Market Price</font></b></td>
	<td width="34%"><b><font color="#ff3333">Transfer Fee</font></b></td>
	<td width="34%"><b><font color="#ff3333">Address</font></b></td>
	<td width="34%"><b><font color="#ff3333">Type</font></b></td>
  </tr>
<%		Do Until rs.EOF%>
                <tr>
                        <td width="34%"><b><font color="#000080"><a href = 'vbScript:ItemSelected(<%=rs.Fields("Security_DPA_")%>, "<%=rs.Fields("SecurityName")%>")'><%=rs.Fields("SecurityName")%></a></font></b></td>
                        <td width="34%"><%=rs.Fields("SecurityCode")%></td>
                        <td width="34%"><%=FormatCurrency(rs.Fields("SecurityMktPrice"),2,,,true)%></td>
                        <td width="34%"><%=FormatCurrency(rs.Fields("SecTransFeeFee"),2,,,true)%></td>
                        <td width="34%"><%=rs.Fields("SecurityAddr")%></td>
                        <td width="34%"><%=rs.Fields("OrderSecTypeDescription")%></td>
				  </tr>
                <%rs.MoveNext
        Loop
        conn.Close
        Set conn = Nothing%>
</table>
    		<input type = 'hidden' name ='ID' id = 'ID' >
    		<input type = 'hidden' name ='action' id = 'action' value="Execute">
</form>




</td>
</tr>
</table>
</div>
</div>

</body>

</html>
