<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "OfferingList"
		const DataEntity = "Offerings"
		const EntityName = "Offering"
		const DataEntityPlural = "Offeringss"
		const ActionFolder = "Operations"
		const Entity_DPA_ = "Offering_DPA_"
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
		
	'Do not allow deleting it has already been batched
    
	sqlstr = "SELECT  Offering_DPA_, Downloaded " & _
			 " FROM   dbo.Offerings " & _
			 " WHERE  (NOT (Batch_No IS NULL)) AND (Offering_DPA_ = " & ID & ")"
     
	 
	 set rs = conn.execute(sqlstr)

	 if not (rs.bof or rs.eof) then
      %>
		<script language="javascript">
		alert('The record cannot be deleted! It has already been batched.');
		</script>
		<%
		set conn = nothing
		Response.end
	 end if	
		
	
		'get voucher number
		sqlStr = "SELECT * FROM Offerings WHERE Offering_DPA_ = " & ID
		
		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If   (rs.BOF Or rs.EOF) Then%>
		        <script language = 'vbscript'>
		        		ShowMessage "Serious error. The payment cannot be found for deletion"
		        		window.self.close
		        </script>
		        <%response.end
		End If
				
		if isnull(rs.fields("Receipt")) then
				 Payment = 0
		else
				Payment = rs.fields("Receipt")
				'voucherType = 3
		end if
		
		clientID = rs.fields("Client_DPA_")						              
	       
		conn.BeginTrans
		'Conn.execute(sqlStr)
		
		'delete from Offerings
		sqlStr = "update [" & DataEntity & "] set deleted=1 WHERE " & Entity_DPA_ & " = " & ID       
         
	
        sqlStr = SQLServerFormat(HandleQuote(sqlStr))	
		
		Conn.execute(sqlStr)	
	
		'conn.execute ("Exec ClientTotalProcedure " & clientID)							
		'conn.execute ("Exec ClientBalanceProcedure " & clientID)
											
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
