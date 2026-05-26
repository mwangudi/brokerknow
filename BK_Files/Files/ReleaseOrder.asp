<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "ReleaseOrder"
	const DataEntity = "Order"
	const DataEntityPlural = "Orders"
	const ActionFolder = "Operations"
	
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("delAction"))
	sourcePage = Request.QueryString("sourcePage")

	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If
        
   	select case action 
		case "EXECUTE"
			Dim hold
			dim ReleaseDate
	        
			hold = Request.Form("HoldStatus")
			ReleaseDate = trim(Request.Form("ReleaseDate"))
			
			'validate Hold
			If Trim(hold) = "" Then%>
					<script language = 'vbscript'>
			    			MsgBox "Invalid hold status"
			        			
					</script>
					<% response.end
			End If
		
			Set conn = GetActiveConnection("KBroker")
	        
			'save data
			Dim holdType
			Dim userID
			Dim manualReleaseDate
			
			if hold = "0" then
					holdType = 1 'release order
			else
					if ReleaseDate = "" then
							holdType = 2 'manual hold
					else
							holdType = 4 'timed hold
					end if
			end if
			
			if holdType = 1 then
					userID = Session("UserID")
					manualReleaseDate = "GETDATE()"
			else	
					userID = "NULL"
					manualReleaseDate = "NULL"
			end if
			
			sqlStr = "UPDATE [tbOrder] SET OrderHold = " & " " & hold & " " & "" & _
					",OrderHoldType_DPA_ = " & holdType & _
					",OrderReleasedBy = " & userID & _
					",OrderDateReleased = " & manualReleaseDate & _
					" WHERE Order_DPA_  = " & ID                

			        
			conn.BeginTrans
					conn.Execute SQLServerFormat(HandleQuote(sqlStr))
			conn.CommitTrans
			conn.Close
			Set conn = Nothing
			
			if (sourcePage = "" or isnull(sourcePage)) then
			    response.redirect "OrderReleaseList.asp"	'Default redirect page    
			else
				response.redirect sourcePage
			end if
    end select
    	
%>

