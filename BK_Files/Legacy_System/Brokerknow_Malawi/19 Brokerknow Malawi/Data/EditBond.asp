<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Bond</title>

<SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

<!--CALENDAR -->
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
<script language="JavaScript" src="CALENDAR/calendar.js"></script>


<!--END CALENDAR -->
</head>

<body Class="Dialog">
<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frmEditSecurity", "txtADate","cmdDate","<%= FormatDate(Date) %>",1);
</SCRIPT>

<script language='vbscript'>
 function ItemSelected(itemID, itemText)
 		Dim ans
 		ans = MsgBox("Delete " & itemText & "?", vbYesNo)
 		if ans = vbNo Then
 			exit function
 		end if
 		frmEditSecurity.elements("ItemID").value = itemID
 		frmEditSecurity.elements("action").value = "Execute_Delete"
 		window.parent.frames("dialogFrame").UpdateDialogHandle
 		frmEditSecurity.submit
 end function
 
 function AddItem()
 		frmEditSecurity.elements("action").value = "Execute_Detail"
 		window.parent.frames("dialogFrame").UpdateDialogHandle	
 		frmEditSecurity.submit
 end function
 </script>

<script language="javascript">
		function  UpdateImmobilised(theChk)
		{
			if (theChk.checked)
			{
				document.frmMain.elements("txtImmobilised").value = "1";
			}
			else
			{
				document.frmMain.elements("txtImmobilised").value = "0";
			}
				
			
		}
</script>

<%
	const LinkedIndependent = 1
   const LinkedDependent = 2
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If

	select case action 
	case "EXECUTE_HEADER"
		Dim name
		Dim addr
		Dim transFee
		Dim mktPrice
		Dim code
		Dim secType
		Dim immob
       
       immob = Request.Form("txtImmobilised")		        
       code = Request.Form("txtCode")
       name = Request.Form("txtName")
		addr = Request.Form("txtAddr")
       mktPrice = Request.Form("txtMktPrice")
       'secType =  Request.Form("cboSecType")
       secType =  1
                    
       
        'validate Name
        If Trim(Name) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Name"
                						
								</script>
				<% response.end
        End If
        
        'validate Market Price
        If Trim(mktPrice) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Market Price"
                						
								</script>
				<% response.end
        End If
        'validate size of Address
        If Len(Addr) > 100 Then%>
				<script language="vbscript">
								ShowMessage "Address can only be 100 characters in length"
								
								</script>
				<% response.end
        End If
        'validate size of Code
        If Len(Code) > 10 Then%>
				<script language="vbscript">
								ShowMessage "Code can only be 10 characters in length"
								
								</script>
				<% response.end
						End If
        'ensure Market Price is numeric
        If (MktPrice <> "") And (Not IsNumeric(MktPrice)) Then%>
			<script language="vbscript">
							ShowMessage "Market Price  must be numeric"
							
							</script>
			<% response.end
        End If
        'validate size of Name
        If Len(Name) > 100 Then%>
			<script language="vbscript">
							ShowMessage "Name can only be 100 characters in length"
							
							</script>
			<% response.end
        End If
        'validate Security Type
        If Trim(secType) = "" Then%>
				<script language="vbscript">
                						ShowMessage "Please specify the Security Type"
                						
								</script>
				<% response.end
        End If
        
        Set conn = GetActiveConnection("KBroker")
      
		 
        'save data
        sqlStr = "UPDATE [Security] SET SecurityAddr = " & "'" & Addr & "'" & ",SecurityCode = " & "'" & Code & "'" & "" & _
                "       ,SecurityMktPrice = " & " " & MktPrice & " " & ",SecurityName = " & "'" & Name & "'" & "" & _
                "       ,OrderSecType_DPA_ = " & " " & secType & _
                "       ,Immobilised = " & " " & immob & _
                "        WHERE Security_DPA_  = " & ID
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
       
        conn.Close
        Set conn = Nothing        
        WritefraEnabledDialogCloseScript
		Response.End
   	case "EXECUTE_DETAIL"
   		Dim adate
       Dim fee		
        
       adate = Request.Form("txtADate")
       fee = Request.Form("txtFee")              
       
        'validate detail info
                'validate Transfer Fee
                If Trim(Fee) = "" Then%>
                		<script language = 'vbscript'>
                        ShowMessage "Please specify the Fee"
                        
                		</script>
                		<% response.end
                End If
                'ensure Transfer Fee is numeric
                If (Fee <> "") And (Not IsNumeric(Fee)) Then%>
                		<script language = 'vbscript'>
                        ShowMessage "Fee must be numeric"
                        
                		</script>
                		<% response.end
                End If

                                        
        'save detail data
         sqlStr = "INSERT INTO [SecTransFee] (SecTransFeeADate,SecTransFeeFee,SecTransFee_DPA_" & _
                "                       ,Security_DPA_) SELECT " & "#" & adate  & " " & Time & "#" & " as SecTransFeeADate" & _
                "       ," & " " & fee & " " & " as SecTransFeeFee" & _
                "       ," & " " & "iif(isnull(max([SecTransFee_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'SecTransFee'),max([SecTransFee_DPA_]) + 1)" & " " & " as SecTransFee_DPA_" & _
                "       ," & " " & ID & " " & " as Security_DPA_" & _
                "        FROM [SecTransFee]"
                                
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
                conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        
        WriteDialogRelocateScript "EditSecurity.asp?ID=" & ID
        'WriteRefreshDialogScript
        Response.End
    case "EXECUTE_DELETE"
    	ItemID = Request.Form("ItemID")

		If Trim(ItemID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for deletion"
                		
                </script>
                <%response.end
        End If
		
		'ensure at least one detail record is left over	
		sqlStr = "SELECT COUNT(SecTransFee_DPA_) as Total FROM [SecTransFee] WHERE Security_DPA_=" & ID
		 Set conn = GetActiveConnection("KBroker")
        
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The database is corrupted"
                		
                </script>
                <%response.end
        End If
        If (CInt(rs.Fields("Total")) < 2) Then%>
                <script language = 'vbscript'>
                		ShowMessage "There must be at least one transfer fee"
                		
                </script>
                <%response.end
        End If
        
        'find out whether any child records exist
        sqlStr = "SELECT Child,DeletionMessage,ParentKey FROM [_Parent_Child_Links_] WHERE (Parent = 'SecTransFee') AND (ChildType = " & LinkedIndependent & ")"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.BOF Or rs.EOF) Then
                Dim childRS
                Dim tableName
                
                rs.MoveFirst
                Do Until rs.EOF
                			tableName = rs.Fields("Child")
                        sqlStr = "SELECT TOP 1 * FROM [" & tableName & "] WHERE SecTransFee_DPA_ = " & ItemID
                        Set childRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                        If Not (childRS.BOF Or childRS.EOF) Then%>
                				<script language = 'vbscript'>
                					ShowMessage "<%=rs.Fields("DeletionMessage")%>"
                					
                				</script>
                				<%response.end
                        End If
                        rs.MoveNext
                Loop
        End If
        
        'delete from database
        sqlStr = "DELETE FROM [SecTransFee] WHERE SecTransFee_DPA_ = " & ItemID
        conn.Execute SQLServerFormat(HandleQuote(sqlStr))
		
		Set Conn = Nothing
		
		WriteDialogRelocateScript "EditSecurity.asp?ID=" & ID
		Response.End
		
   	end select
%>
<form name = 'frmEditSecurity' method = 'post' id="frmMain" action = 'EditBond.asp' >
<table border="0" width="100%">
  <tr><%
		
		Set conn = GetActiveConnection("KBroker")
        
        
		sqlStr = "SELECT Immobilised,SecurityAddr,SecurityCode,SecurityMktPrice,SecurityName" & _
                "       ,OrderSecType_DPA_,Security_DPA_ FROM [Security] WHERE Security_DPA_  = " & ID
        
        sqlStr = SQLServerFormat(HandleQuote(sqlStr))
        Set rs = conn.Execute(sqlStr)
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Security cannot be retrieved for editing"
                		
                </script>
                <% response.end
        End If
       
        %>

    <td width="30%">Name</td>
    <td width="70%"><input type = 'text' name ='txtName' id = 'txtName' value = '<%=rs.Fields("SecurityName")%>' size="30"></td>
  </tr>
  <tr>
  <td width="30%">Code</td>
    <td width="70%"><input type = 'text' name ='txtCode' id = 'txtCode' value = '<%=rs.Fields("SecurityCode")%>' size="30"></td>
  </tr>
  <tr>
  <td width="30%">Address</td>
    <td width="70%"><textarea rows=3 name ='txtAddr' id = "txtAddr"><%=rs.Fields("SecurityAddr")%></textarea>
    </td>
  </tr>
  <tr>
  <td width="30%">Market Price</td>
    <td width="70%"><input type = 'text' name ='txtMktPrice' id = 'txtMktPrice' value = '<%=rs.Fields("SecurityMktPrice")%>' size="30"></td>
  </tr>
  
  <tr>
    <td>Immobilised</td>
    <td>
		<%	Dim chkState
			Dim applicable
    
			if rs.Fields("Immobilised") then
				chkState = "checked"
				applicable = "1"
			else
				chkState = ""
				applicable = "0"
			end if
		%>
			<input type=checkbox Class="BorderLess"   <%=chkState%> name='chkImmobilised' onClick = 'UpdateImmobilised(this);'> 
			<input type = 'hidden' name ='txtImmobilised' id = 'txtImmobilised' value='<%=applicable%>'>
      </td>
  </tr>
  <tr>
    <td width="30%"><input type = 'button' Class=Buttons OnClick="VBScript: AddItem" name ='cmdAdd' id = 'cmdAdd' value="Add Transfer Fee"></td>
    <td width="70%">

    </td>
  </tr>
  
  <tr>
    <td colspan = '2'>
    <table border="0" width="100%">
    <tr>
      <td width="32%"><b><font color="#000080">Transfer Fee</font></b></td>
      <td width="32%"><b><font color="#000080">Activation Date</font></b></td>

     </tr>
    <tr>
    
      <td width="24%"><input type = 'text' name ='txtFee' id = 'txtFee' size="20">
      
    	</td>
    	<td width="24%"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
    </tr>
 <%
	 sqlStr = "SELECT SecTransFeeADate,SecTransFeeFee,SecTransFee_DPA_" & _
                "        FROM [SecTransFee] WHERE Security_DPA_  = " & ID & " ORDER BY SecTransFeeADate DESC"	
    Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If Not(rs.EOF Or rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF%>
        		<tr>
                       <td><a href = 'vbScript:ItemSelected(<%=rs.Fields("SecTransFee_DPA_")%>, "<%=rs.Fields("SecTransFeeFee")%>")'><%=rs.Fields("SecTransFeeFee")%></a></td>
                       <td><%= FormatDate(rs.Fields("SecTransFeeADate")) %></td>
                        
             </tr>
             <%rs.MoveNext
        Loop
   End if
 %>
  </table>
    </td>
  </tr>
  <tr>
   <td width="100%" colspan="2" align=right>
		<BR>	
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		<input type = 'button' Class=Buttons name ='cmdCancel' id = 'cmdCancel' value=" Cancel " OnClick="JavaScript: window.self.close();">
		&nbsp;&nbsp;
		<input type = 'hidden' name ='action' id = 'action' value="Execute_Header">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
    	<input type = 'hidden' name ='ItemID' id = 'ItemID'>
      </td>
  </tr>
</table>
</form>


</body>

</html>
