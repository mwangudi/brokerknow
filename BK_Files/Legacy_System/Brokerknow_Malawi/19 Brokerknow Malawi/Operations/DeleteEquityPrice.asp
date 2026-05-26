<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Delete Market Price</title>
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
 <script language="JavaScript" src="../scripts/common.js"></script>
</head>

<body>
<!--#include file="../libroutines.asp"-->

<%

'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "EquityList"
		const DataEntity = "MarketPrice"
		const EntityName = "Security"
		const DataEntityPlural = "Market Prices"
		const ActionFolder = "Operations"
'======================= End_Alter_Across_Entities =================================		
%>

<%

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
			    
			    'Retrieve Security_DPA_ AND Market Date from the passed ID value

					Security_DPA = mid(ID,1,instr(1,ID,"-")-1)
					MktDate = mid(ID,instr(1,ID,"-")+1,len(ID)-1)

				'Retrieve security
				
				sqlStr = "SELECT * FROM " & DataSource & "  WHERE Security_DPA_ = " & Security_DPA & " AND [Date] = #" & MktDate & "#" 
				
				Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				
				If   (rs.BOF Or rs.EOF) Then%>
				        <script language = 'vbscript'>
				        		ShowMessage "Serious error. The Security marke price cannot be found for deletion"
				        		window.self.close
				        </script>
				        <%response.end
				End If
				
				if isnull(rs.fields("MktUnique")) then
					MktUnique = 0	 
				else
					MktUnique = rs.fields("MktUnique")
				end if
			

				'delete from database
				sqlStr = "DELETE FROM datastream_Market WHERE MktUnique = " & MktUnique 
				 
				sqlStr = SQLServerFormat(HandleQuote(sqlStr))

				conn.BeginTrans 
					 conn.execute sqlStr	
				conn.CommitTrans
				
				WritefraEnabledDialogCloseScript
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

</body>

</html>
