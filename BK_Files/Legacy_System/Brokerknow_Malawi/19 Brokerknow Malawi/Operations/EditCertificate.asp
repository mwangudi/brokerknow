<!--#include file="../libroutines.asp"-->
<%
	
	'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "EditCertificate"
		const DataEntity = "Certificate"
		const DataEntityPlural = "Certificates"
		const ActionFolder = "Operations"
'======================= End_Alter_Across_Entities =================================
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim ID
	
	action = ucase(Request.Form("action"))
	
	ID = Request("ID")
	If (Trim(ID) = "") or (ID = "0") Then%>
			<script language = 'vbscript'>
                	ShowMessage "Please select a certificate for editing"
                	window.self.close
			</script>
			<%response.end
	End If
	
	response.Redirect "Add" & DataEntity & ".asp?ID=" & ID
%>