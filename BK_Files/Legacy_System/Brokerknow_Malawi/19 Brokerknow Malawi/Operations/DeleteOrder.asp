<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Delete Order</title>


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
	UserId=Session("UserId")
	Dim conn 
   Dim sqlStr
   Dim rs
	
	Set conn = GetActiveConnection("KBroker")
    
        
	action = ucase(Request.Form("delAction"))
	if action = "EXECUTE" then
		  
       ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for deletion"
                		
                </script>
                <%response.end
        End If
		
		'convert item ID to order ID
		Dim orderRS
		sqlStr = "SELECT OrdDetailList.Order_DPA_, OrdDetailList.OrderHold FROM OrdDetailList WHERE  OrdDetailList.OrdDetail_DPA_=" & ID 
		Set orderRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If (orderRS.EOF Or orderRS.BOF) Then%>
				    <script language = 'vbscript'>
				         	ShowMessage "The Order cannot be retrieved for deletion"
				         	
				    </script>
				    <% response.end
		End If
		ID = orderRS.Fields("Order_DPA_")
        'find out whether any child records exist
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'tbOrder') AND (ChildType = " & LinkedIndependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                Dim depRS
                Dim tableName
                
                rs.MoveFirst
                Do Until rs.EOF
                		tableName = rs.Fields("Child")
                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE Order_DPA_ = " & ID & " and Deleted=0"
                        Set depRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                        If Not (depRS.BOF Or depRS.EOF) Then%>
                				<script language = 'vbscript'>
                					ShowMessage "<%=rs.Fields("DeletionMessage")%>"
                					
                				</script>
                				<%response.end
                        End If
                        rs.MoveNext
                Loop
        End If
        
        'check whether child records are linked to other data
        Dim childRS
        Dim childName
        
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'tbOrder') AND (ChildType = " & LinkedDependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                rs.MoveFirst
                childName = rs.Fields("Child")
                sqlStr = "SELECT * FROM [" & childName & "] WHERE Order_DPA_ = " & ID & " and Deleted=0"
                Set childRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                If Not (childRS.BOF Or childRS.EOF) Then
                        'find the key field
                        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = '" & childName & "') AND (ChildType = " & LinkedIndependent & ")"
                        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                        If Not (rs.BOF Or rs.EOF) Then
                                Dim childKey
                                
                                childKey = rs.Fields("ParentKey")
                                Do Until childRS.EOF
                                        'find out whether any child records exist
                                        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = '" & childName & "') AND (ChildType = " & LinkedIndependent & ")"
                                        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                                        If Not (rs.BOF Or rs.EOF) Then
                                                rs.MoveFirst
                                                
                                                Do Until rs.EOF
														tableName = rs.Fields("Child")
                                                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE " & childKey & " = " & childRS.Fields(childKey) & " and Deleted=0"
                                                        Set depRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                                                        If Not (depRS.BOF Or depRS.EOF) Then%>
                														<script language = 'vbscript'>
                																ShowMessage "<%=rs.Fields("DeletionMessage")%>"
                																
                														</script>
                														<%response.end
                                                        End If
                                                        rs.MoveNext
                                                Loop
                                        End If
                                        childRS.MoveNext
                                Loop
                                childRS.MoveFirst
                        End If
                End If
        End If
        
        'delete child records
        conn.BeginTrans
                'delete from database
                if cbool(orderRS.fields("OrderHold")) then
						If Not (childRS.BOF Or childRS.EOF) Then
						        sqlStr = "Update [" & childName & "] Set Deleted=1 WHERE Order_DPA_ = " & ID
						        conn.Execute SQLServerFormat(HandleQuote(sqlStr))
						End If
						sqlStr = "Update [tbOrder] Set Deleted=1,ChangedBy= " & UserId & ",TimeChanged=GetDate() WHERE Order_DPA_ = " & ID
						Conn.Execute("Update OrdDetail set Deleted=1 WHERE Order_DPA_ = " & ID )
				else
						sqlStr = "UPDATE [tbOrder] SET Deleted=1,OrderCanceled = 1,ChangedBy="& UserId &",TimeChanged=GetDate() WHERE Order_DPA_ = " & ID						
						Conn.Execute("Update OrdDetail set Deleted=1 WHERE Order_DPA_ = " & ID )
				end if
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        Set Conn = Nothing
        WriteDeleteCloseScript
        Response.End
        
   	end If

		 sqlStr = "SELECT * FROM [OrderList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then
                %><p>No Orders found</p><%
                response.end
        End If
        
        rs.MoveFirst
        
%>
<form name = 'frmDeleteOrderList' method = 'post' action = 'DeleteOrderList.asp' >
<table border="0" width="100%">
<tr>
	<td width="34%"><b><font color="#ff3333">Client</font></b></td>
	<td width="34%"><b><font color="#ff3333">Date</font></b></td>
	<td width="34%"><b><font color="#ff3333">Reference No.</font></b></td>
	<td width="34%"><b><font color="#ff3333">Type</font></b></td>
  </tr>
<%		Do Until rs.EOF%>
                <tr>
                        <td width="34%"><%=rs.Fields("ClientName")%></td>
                        <td width="34%"><%=rs.Fields("OrderDate")%></td>
                        <td width="34%"><b><font color="#000080"><a href = 'vbScript:ItemSelected(<%=rs.Fields("Order_DPA_")%>, "<%=rs.Fields("OrderRef")%>")'><%=rs.Fields("OrderRef")%></a></font></b></td>
                        <td width="34%"><%=rs.Fields("OrderTypeName")%></td>
				  </tr>
                <%rs.MoveNext
        Loop
        conn.Close
        Set conn = Nothing%>
</table>
    		<input type = 'hidden' name ='ID' id = 'ID' >
    		<input type = 'hidden' name ='action' id = 'action' value="Execute">
</form>

</body>

</html>
