<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "PaymentRequest"
		const DataEntity = "PaymentRequest"
		const EntityName = "PaymentRequests"
		const DataEntityPlural = "PaymentRequests"
		const ActionFolder = "Operations"
		const Entity_DPA_ = "Request_DPA_"
'======================= End_Alter_Across_Entities =================================		


	const LinkedIndependent = 1
   const LinkedDependent = 2
	
	Dim conn 
   Dim sqlStr
   Dim rs
	
	Set conn = GetActiveConnection("KBroker")
    
        
	action = ucase(Request("delAction"))
	Dim reloadRequired
		
	reloadRequired = false
	if action = "EXECUTE" then
		
	   ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for deletion"
                		window.self.close
                </script>
                <%response.end
        End If		
		
		
	   sqlStr = "SELECT * FROM PaymentRequests WHERE Request_DPA_ = " & ID
		
		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If   (rs.BOF Or rs.EOF) Then%>
		        <script language = 'vbscript'>
		        		ShowMessage "Serious error. The Request cannot be found for deletion"
		        		window.self.close
		        </script>
		        <%response.end
		End If      
		
		clientID = rs.fields("Client_DPA_")
		       
		conn.BeginTrans
		'Conn.execute(sqlStr)
		
		'delete from Offerings	

		sqlStr = "update [" & DataEntity & "] set deleted=1 WHERE " & Entity_DPA_ & " = " & ID & " and Approved=0"
        sqlStr = SQLServerFormat(HandleQuote(sqlStr))
		
		Conn.execute(sqlStr)
		
		conn.execute ("Exec ClientTotalProcedure " & clientID)							
		conn.execute ("Exec ClientBalanceProcedure " & clientID)
											
		conn.CommitTrans
		
        WriteDeleteCloseScript
        Response.End
             
   	end If
   	Dim clientCode
        
    clientCode = "var validNavigate = true;" & chr(13)
	clientCode = clientCode & "window.self.close();" & chr(13)%>
	<script>
		<%=clientCode%>
	</script>
	<%
	response.End
        
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Delete <%=EntityName%></title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <script language="JavaScript" src="../scripts/common.js"></script>
</head>

<body>
</body>

</html>
