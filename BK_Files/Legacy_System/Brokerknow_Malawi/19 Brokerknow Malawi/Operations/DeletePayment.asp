<!--#include file="../libroutines.asp"-->
<%
'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "PaymentList"
		const DataEntity = "Payment"
		const EntityName = "Payment"
		const DataEntityPlural = "Payments"
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
				sqlStr = "SELECT * FROM Payment WHERE Payment_DPA_ = " & ID
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				If   (rs.BOF Or rs.EOF) Then%>
				        <script language = 'vbscript'>
				        		ShowMessage "Serious error. The payment cannot be found for deletion"
				        		window.self.close
				        </script>
				        <%response.end
				End If
				
				if isnull(rs.fields("Voucher_DPA_")) then
						 if isnull(rs.fields("ClientVoucher_DPA_")) then
								 voucher = 0
						else
								voucher = rs.fields("ClientVoucher_DPA_")
								voucherType = 1 'Client
						end if
				else
						voucher = rs.fields("Voucher_DPA_")
						voucherType = 3 'Broker
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

				'delete from database
				sqlStr = "DELETE FROM [" & DataEntity & "] WHERE " & Entity_DPA_ & " = " & ID
				sqlStr = SQLServerFormat(HandleQuote(sqlStr))

				 conn.BeginTrans
						conn.Execute sqlStr
						if voucher <> 0 then
								if (voucherType = 1) then
										'find out whether any child records exist
										sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'ClientVoucher') AND (ChildType = " & LinkedIndependent & ")"
										Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
										If Not (rs.BOF Or rs.EOF) Then
										        rs.MoveFirst
										        Do Until rs.EOF
										        		tableName = rs.Fields("Child")
										                sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE ClientVoucher_DPA_ = " & voucher
										                Set depRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
										                If Not (depRS.BOF Or depRS.EOF) Then%>
										        				<script language = 'vbscript'>
										        					ShowMessage <%=rs.Fields("DeletionMessage")%>
												        					
										        				</script>
										        				<%response.end
										                End If
										                rs.MoveNext
										        Loop
										End If
										'remove contracts 
										
										sqlStr = "UPDATE Contract SET ClientVoucher_DPA_ = NULL" & _
												", ContractClientVouchered = 0 WHERE ClientVoucher_DPA_ =" & voucher
										
										conn.Execute SQLServerFormat(HandleQuote(sqlStr))
										
										'delete from database
										sqlStr = "DELETE FROM [ClientVoucher] WHERE ClientVoucher_DPA_ = " & voucher
										
										conn.Execute SQLServerFormat(HandleQuote(sqlStr))
							
								elseif (voucherType = 3) then
										'find out whether any child records exist
										sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'Voucher') AND (ChildType = " & LinkedIndependent & ")"
										Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
										If Not (rs.BOF Or rs.EOF) Then
										        rs.MoveFirst
										        Do Until rs.EOF
										        		tableName = rs.Fields("Child")
										                sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE Voucher_DPA_ = " & voucher
										                Set depRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
										                If Not (depRS.BOF Or depRS.EOF) Then%>
										        				<script language = 'vbscript'>
										        					ShowMessage <%=rs.Fields("DeletionMessage")%>
										        					
										        				</script>
										        				<%response.end
										                End If
										                rs.MoveNext
										        Loop
										End If
										'remove contracts 
										sqlStr = "UPDATE Contract SET Voucher_DPA_ = NULL" & _
												", ContractVouchered = 0 WHERE Voucher_DPA_ =" & voucher
										conn.Execute SQLServerFormat(HandleQuote(sqlStr))
										'delete from database
										sqlStr = "DELETE FROM [Voucher] WHERE Voucher_DPA_ = " & voucher
										conn.Execute SQLServerFormat(HandleQuote(sqlStr))
								end if
						end if
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
