<!--#include file="../libroutines.asp"-->
<%
	
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
	
	ID = Request.Form("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		MsgBox "No client specified for activation"
                		
                </script>
                <% response.end
        End If
		
		Dim CDS
		        
		CDS = Request.Form("CDSNo")		
		
		'validate CDS number
		If Trim(CDS) = "" Then%>
				<script language = 'vbscript'>
                		MsgBox "Please specify the CDS number"
                		
				</script>
				<% response.end
		End If
		
		'validate size of CDS number
		If Len(CDS) > 20 Then%>
				<script language = 'vbscript'>
				alert "CDS number can only be 20 characters in length"
						
				</script>
				<% response.end
		End If
				
		Set conn = GetActiveConnection("KBroker")
		
		'save data
		sqlStr = "UPDATE [Client] SET ClientCDSNo = " & "'" & CDS & "'" & "" & _
				" WHERE Client_DPA_  = " & ID
			    
		
		conn.BeginTrans
				conn.Execute SQLServerFormat(HandleQuote(sqlStr))
		conn.CommitTrans
		conn.Close
		Set conn = Nothing

		response.redirect "ClientActivationList.asp"
%>
















