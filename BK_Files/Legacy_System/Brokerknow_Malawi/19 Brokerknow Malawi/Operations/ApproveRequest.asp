<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "ApproveRequest"
	const DataEntity = "PaymentRequest"
	const DataEntityPlural = "PaymentRequests"
	const ActionFolder = "Operations"
	
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("delAction"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If
        
        
	'Ids = Split(ID,"<->")
	
	'ClientID = Ids(1)
	'itemID =Ids(0)
	
	'Response.Write
	'Response.End
	
	select case action 
		case "EXECUTE"
			Dim release
			dim ReleaseDate
	        
			release = Request.Form("Release")
			ReleaseDate = trim(Request.Form("ReleaseDate"))
			
			'validate release
			If Trim(release) = "" Then%>
					<script language = 'vbscript'>
			    			MsgBox "Invalid release status"
			        			
					</script>
					<% response.end
			End If
		
			Set conn = GetActiveConnection("KBroker")
	        
			'save data
			Dim userID
			Dim manualReleaseDate
			
			manualReleaseDate = "GETDATE()"
			
			sqlStr = "UPDATE PaymentRequests SET Approved = " & " " & release & " " & "" & _
					",ApprovedBy="& session("Userid") &",ApprovalDate = " & manualReleaseDate & _
					" WHERE Processed_DPA_  in (" & request.form("chkReleaseJournal") &")"
		
		'response.write sqlStr
		'response.end
		
			conn.BeginTrans
					conn.Execute SQLServerFormat(HandleQuote(sqlStr))
			conn.CommitTrans
			conn.Close
			Set conn = Nothing

			response.redirect "ApproveRequests.asp"	
    end select
    	
%>

