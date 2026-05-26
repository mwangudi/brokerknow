<!--#include file="../libroutines.asp"-->
<%
	
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
	
	ID = Request.Form("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% response.end
        End If
		
		Dim name
		Dim addr
		Dim cell
		Dim email
		Dim contact
		Dim office
				
		name = Request.Form("Name")
		addr = Request.Form("Address")
		cell = Request.Form("CellPhone")
		contact = Request.Form("Contact")
		email = Request.Form("Email")
		office = Request.Form("OfficePhone")
		
		'validate Name
		If Trim(Name) = "" Then%>
				<script language = 'vbscript'>
                		ShowMessage "Please specify the Name"
                		
				</script>
				<% response.end
		End If
		'validate size of Address
		If Len(Addr) > 100 Then%>
				<script language = 'vbscript'>
				ShowMessage "Address can only be 100 characters in length"
				
				</script>
				<% response.end
		End If
		'validate size of Cell Phone
		If Len(Cell) > 100 Then%>
				<script language = 'vbscript'>
				ShowMessage "Cell Phone can only be 100 characters in length"
				
				</script>
				<% response.end
		End If
		'validate size of Contact Name
		If Len(Contact) > 100 Then%>
				<script language = 'vbscript'>
				ShowMessage "Contact Name can only be 100 characters in length"
				
				</script>
				<% response.end
		End If
		'validate size of Email
		If Len(Email) > 100 Then%>
				<script language = 'vbscript'>
				ShowMessage "Email can only be 100 characters in length"
				
				</script>
				<% response.end
		End If
		'validate size of Name
		If Len(Name) > 100 Then%>
				<script language = 'vbscript'>
				ShowMessage "Name can only be 100 characters in length"
				
				</script>
				<% response.end
		End If
		'validate size of Office Phone
		If Len(Office) > 100 Then%>
				<script language = 'vbscript'>
				ShowMessage "Office Phone can only be 100 characters in length"
				
				</script>
				<% response.end
		End If
		
		Set conn = GetActiveConnection("KBroker")
		
		'save data
		sqlStr = "UPDATE [Client] SET ClientAddr = " & "'" & addr & "'" & "" & _
				",ClientCellTel = " & "'" & cell & "'" & ",ClientContact = " & "'" & contact & "'" & "" & _
				",ClientEmail = " & "'" & email & "'" &  "" & _
				",ClientName = " & "'" & name & "'" & ",ClientOfficeTel = " & "'" & Office & "'" & "" & _
				" WHERE Client_DPA_  = " & ID
			    
		
		conn.BeginTrans
				conn.Execute SQLServerFormat(HandleQuote(sqlStr))
		conn.CommitTrans
		conn.Close
		Set conn = Nothing

		response.redirect "ClientList.asp"
%>
















