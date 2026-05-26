<!--#include file="../libroutines.asp"-->
<%
	
	'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "EditCertificate"
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
                	ShowMessage "Please select a Contract"
                	window.self.close
			</script>
			<%response.end
	End If
	
	Dim certificate
	Dim rDate
    
    certificate = Request.Form("NewCert")
    rDate = Request.Form("RDate")
    
    
    'validate Certificate
    If Trim(certificate) = "" Then%>
            <script language = 'vbscript'>
                	ShowMessage "Please specify the Certificate"
                	
            </script>
            <% response.end
    End If
    'validate size of Certificate
    If Len(certificate) > 20 Then%>
            <script language = 'vbscript'>
            ShowMessage "Certificate can only be 20 characters in length"
            
            </script>
            <% response.end
    End If
    
    'save data
    sqlStr = "UPDATE [Contract] SET ContractNCertificate = '" & certificate &  "'," & _
                " ContractNCDate = #" & FormatDate(rDate) & "#, " & _
                " ContractNCDelivered = 1 WHERE Contract_DPA_ = " & ID
    Set conn = GetActiveConnection("KBroker")
    
    conn.BeginTrans
            conn.Execute SQLServerFormat(HandleQuote(sqlStr))
    conn.CommitTrans
    conn.Close
    Set conn = Nothing
    response.redirect DataEntity & "List.asp"
    Response.End
   %>