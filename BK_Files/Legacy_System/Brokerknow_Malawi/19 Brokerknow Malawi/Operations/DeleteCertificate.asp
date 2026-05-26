<!--#include file="../libroutines.asp"-->
<%
	
	'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "DeleteCertificate"
		const DataEntity = "Certificate"
		const DataEntityPlural = "Certificates"
		const ActionFolder = "Operations"
'======================= End_Alter_Across_Entities =================================
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim ID
	
	ID = Request("ID")
	If (Trim(ID) = "") or (ID = "0") Then%>
			<script language = 'vbscript'>
                	ShowMessage "Please select a Certificate for deletion"
                	window.self.close
			</script>
			<%response.end
	End If
	
	'save data
    sqlStr = "UPDATE [Contract] SET ContractNCertificate = ''," & _
            " ContractNCDate = NULL, " & _
            " ContractNCDelivered = 0 WHERE Contract_DPA_ = " & ID
    Set conn = GetActiveConnection("KBroker")
    
    conn.BeginTrans
            conn.Execute SQLServerFormat(HandleQuote(sqlStr))
    conn.CommitTrans
    conn.Close
    Set conn = Nothing
    WriteDeleteCloseScript
    response.end
	%>
	
