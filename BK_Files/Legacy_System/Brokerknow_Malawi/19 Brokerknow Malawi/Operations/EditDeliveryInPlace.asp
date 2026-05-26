<!--#include file="../libroutines.asp"-->
<%
	
	'======================= Begin_Alter_Across_Entities =================================
		
		const UDLName = "KBroker"
		const DataSource = "EditDelivery"
		const DataEntity = "Delivery"
		const DataEntityPlural = "Deliveries"
		const ActionFolder = "Operations"
'======================= End_Alter_Across_Entities =================================
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim ID
	
	ID = Request("ID")
	If (Trim(ID) = "") or (ID = "0") Then%>
			<script language = 'vbscript'>
                	ShowMessage "Please select a Contract for delivery"
                	window.self.close
			</script>
			<%response.end
	End If
	
	Dim transferNo
	Dim dDate
    
    transferNo = Request.Form("Transfer")
    dDate = Request.Form("DDate")
    
    
    'validate Transfer No
    If Trim(transferNo) = "" Then%>
            <script language = 'vbscript'>
                	ShowMessage "Please specify the Transfer No"
                	
            </script>
            <% response.end
    End If
    'validate size of Transfer No
    If Len(transferNo) > 100 Then%>
            <script language = 'vbscript'>
            ShowMessage "Transfer No can only be 20 characters in length"
            
            </script>
            <% response.end
    End If
    
    'save data
    sqlStr = "UPDATE [Contract] SET ContractTransferNo = '" & transferNo &  "'," & _
            " ContractDeliveryDate = #" & FormatDate(dDate) & "#, " & _
            " ContractDelivered = 1 WHERE Contract_DPA_ = " & ID
    Set conn = GetActiveConnection("KBroker")
    
    conn.BeginTrans
            conn.Execute SQLServerFormat(HandleQuote(sqlStr))
    conn.CommitTrans
    conn.Close
    Set conn = Nothing
    response.redirect DataEntity & "List.asp"
    Response.End
   %>