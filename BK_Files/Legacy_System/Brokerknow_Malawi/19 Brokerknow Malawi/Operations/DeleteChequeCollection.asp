<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "ChequeCollection"
		const DataEntity = "ChequeCollection"
		const EntityName = "Cheque Collection"
		const DataEntityPlural = "Cheque Collection"
		const ActionFolder = "Operations"
		const Entity_DPA_ = "Payment_DPA_"
'======================= End_Alter_Across_Entities =================================		


		const LinkedIndependent = 1
		const LinkedDependent = 2
	
		Dim conn 
		Dim sqlStr
		Dim rs
		Dim voucher
		Dim voucherType
	
		Set conn = GetActiveConnection("KBroker")
    
		action = ucase(Request("delAction"))
		Dim reloadRequired
		'STOP
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
			
				conn.BeginTrans
					
					sqlStr = "UPDATE Payment SET ChequeCollection = N'AWAITING COLLECTION' WHERE (Payment_DPA_ = " & ID & ") "
					'sqlStr = SQLServerFormat(HandleQuote(sqlStr))
					conn.execute(sqlStr)
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
