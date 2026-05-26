<!--#include file="../libroutines.asp"-->


<%
	const UDLName = "KBroker"
	const DataSource = "StockWatchList"
	const DataEntity = "StockWatchList"
	const DataEntityPlural = "StockWatchLists"
	const ActionFolder = "Operations"
	const ActionPage = "StockWatchList"

	Dim conn 
	Dim sqlStr
	Dim rs
	
		Set conn = GetActiveConnection("KBroker")
	 
	    UserId=Session("UserID")
	     
		action = ucase(Request.Form("delAction"))
		if action = "EXECUTE" then
			  
	    ID = Request("ID")
		ID1= split(ID,"-")
		itemId= ID1(0)
		clientID =ID1(1)

			If Trim(itemId) = "" Then%>
	             <script language = 'vbscript'>
	             		ShowMessage "No record specified for deletion"
	             		
	             </script>
	             <%response.end
	     End If
			
			
	             'delete from database
				sqlStr = "Update [StockWatch] Set Deleted=1 WHERE StockWatch_DPA_ = " & ItemID
		 conn.BeginTrans	
	            conn.Execute SQLServerFormat(HandleQuote(sqlStr))
	     conn.CommitTrans
	     Set Conn = Nothing
	     WriteDeleteCloseScript
	     Response.End
	     
		end If

%>

