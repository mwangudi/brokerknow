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
				
        'find out whether any child records exist
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = '" & DataEntity & "') AND (ChildType = " & LinkedIndependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                Dim childRS
                Dim tableName
                
                rs.MoveFirst
                Do Until rs.EOF
                			tableName = rs.Fields("Child")
                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE " & Entity_DPA_ & " = " & ID
                        
                        Set childRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                        If Not (childRS.BOF Or childRS.EOF) Then%>
                				<script language = 'vbscript'>
                					ShowMessage "<%=rs.Fields("DeletionMessage")%>"
                					window.self.close
                				</script>
                				<%response.end
                        End If
                        rs.MoveNext
                Loop
        End If       
        
        'Remove From Payment
				
		SqlStr="update Payment set Deleted=1 Where (Payment_DPA_=" & Payment & ")"
		sqlStr = SQLServerFormat(HandleQuote(sqlStr))
		       
		conn.BeginTrans
		Conn.execute(sqlStr)
		
		'delete from Offerings
		sqlStr = "update [" & DataEntity & "] set deleted=1 WHERE " & Entity_DPA_ & " = " & ID
        sqlStr = SQLServerFormat(HandleQuote(sqlStr))
		
		Conn.execute(sqlStr)
				
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
