<!--#include file="../libroutines.asp"-->
<%
	
	'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "EditDelivery"
		const DataEntity = "Delivery"
		const DataEntityPlural = "Deliveries"
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
                	ShowMessage "Please select a delivery for editing"
                	window.self.close
			</script>
			<%response.end
	End If
	
	response.Redirect "EditSDeliveryInPlace.asp?ID=" & ID
%>