<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "SubmitOffering"
	const DataEntity = "Offering"
	const DataEntityPlural = "Offerings"
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
        
        

	select case action 
		case "EXECUTE"
			Dim hold
			dim ReleaseDate
	        
			hold = Request.Form("HoldStatus")
			ReleaseDate = trim(Request.Form("ProposalDate"))
			
			'validate Hold
			If Trim(hold) = "" Then%>
					<script language = 'vbscript'>
			    			MsgBox "Invalid accept status"
			        			
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
			end if
			
			if holdType = 1 then
					userID = Session("UserID")
					manualReleaseDate = "GETDATE()"
			end if
			
			if(Cint(hold)=1) then
			sqlStr = "UPDATE BondProposals SET Accepted = " & " " & hold & " " & "" & _
					",ProposalDate = '" & Date() & "'" & _
					" WHERE Bond_DPA_  = " & ID                
			else
			sqlStr = "UPDATE BondProposals SET Accepted = " & " " & hold & " " & "" & _
					",ProposalDate = null" & _
					" WHERE Bond_DPA_  = " & ID                
			end if        
			'Response.Write(sqlStr)
			'Response.End
			
			conn.BeginTrans
					conn.Execute SQLServerFormat(HandleQuote(sqlStr))
			conn.CommitTrans
			conn.Close
			Set conn = Nothing
			
			if(Cint(Hold)=1) then
			response.redirect "AcceptProposalList.asp"	
			else
			response.redirect "UndoAcceptProposalList.asp"	
			end if
    end select
    	
%>

