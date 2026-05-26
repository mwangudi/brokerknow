<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "EditReceipt"
	const DataEntity = "Receipt"
	const DataEntityPlural = "Receipts"
	const ActionFolder = "Operations"
	
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim rsEdit
	Dim guid
	Dim guidStr
	
	ID = Request("ID")

	If Trim(ID) = "" Then%>
			<script language = 'vbscript'>
					ShowMessage "No record specified for editing"
	                
			</script>
			<% response.end
	End If

	Dim amount
	Dim reference
	
	amount = Request.Form("Amount")
	reference = Request.Form("Reference")
	
	'validate Amount
	If Trim(Amount) = "" Then%>
		<script language = 'vbscript'>
				ShowMessage "Please specify the Amount "
				
		</script>
		<% response.end
	End If
	'ensure Amount is numeric
	If (Amount <> "") And (Not IsNumeric(Amount)) Then%>
		<script language = 'vbscript'>
			ShowMessage "Order Detail Estimated Amount must be numeric"
			
		</script>
		<% response.end
	End If
	
	'validate size of Reference 
		If Len(Reference) > 100 Then%>
				<script language = 'vbscript'>
				ShowMessage "Reference can only be 20 characters in length"
				
				</script>
				<% response.end
		End If

	'save data		
	Set conn = GetActiveConnection("KBroker")
	
	sqlStr = "UPDATE CPayment SET CPaymentAmount = " & amount & _
			", CPaymentReference = '" & reference & "'" & _
			" WHERE CPayment_DPA_= " & ID
	conn.Execute SQLServerFormat(HandleQuote(sqlStr))
	
	conn.Close
	Set conn = Nothing
	response.redirect DataEntity & "List.asp"
%>

