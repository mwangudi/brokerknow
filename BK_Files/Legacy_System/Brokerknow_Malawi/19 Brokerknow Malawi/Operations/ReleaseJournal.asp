<!--#include file="../libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "ReleaseJournal"
	const DataEntity = "Journal"
	const DataEntityPlural = "Journals"
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
			
			sqlStr = "UPDATE Journal SET Released = " & " " & release & " " & "" & _
					",ReleaseDate = " & manualReleaseDate & _
					" WHERE Journal_DPA_  = " & ID                

			conn.BeginTrans
					conn.Execute SQLServerFormat(HandleQuote(sqlStr))
			conn.CommitTrans
			conn.Close
			Set conn = Nothing

			response.redirect "JournalRelease.asp"	
    end select
    	
%>

