<!--#include file="../libroutines.asp"-->

<%
	const UDLName = "KBroker"
	const DataSource = "DeleteLot"
	const DataEntity = "Lot"
	const DataEntityPlural = "Lots"
	const ActionFolder = "Operations"
	
	const LinkedIndependent = 1
    const LinkedDependent = 2
	
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim guidStr 
	Dim guid
	Dim ID
	Dim ItemID
	Dim rsEdit
	Dim IDHolder
	Dim IDArray
	
	UserId=Session("UserID")
	
	IDHolder = Request("ID")
	
	If (Trim(IDHolder) = "") or (IDHolder = "0") Then%>
			<script language = 'vbscript'>
                	ShowMessage "No item selected for deletion"
                	window.self.close
			</script>
			<%response.end
	End If
	IDArray = split(IDHolder,"<->")
	ID = IDArray(lbound(IDArray))
	ItemID = IDArray(ubound(IDArray))	
	
	if itemID <> "0" then
			Set conn = GetActiveConnection("KBroker")
			
			'obtain levies, contract and lot to be deleted
			sqlStr = "SELECT * FROM LevyContractList WHERE Lot_DPA_=" & ItemID
			Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			If (rs.EOF Or rs.BOF) Then%>
						<script language = 'vbscript'>
				         		ShowMessage "No Levies were found for this Lot. This is a serious database corruption."
								
						</script>
						<% response.end
			End If
			Dim contractID
			Dim lotID
			Dim transIsCDS
			
			contractID = rs.fields("Contract_DPA_")
			lotID = rs.fields("Lot_DPA_")

			conn.execute "Update Lot set Changedby = " & UserId & ", Timechanged = Getdate() where contract_dpa_ = " & contractID
			conn.execute "Delete from Lot where contract_dpa_ = " & contractID
			
			conn.Close
			Set conn = Nothing
	end if 
	WriteDeleteCloseScript
    response.end 		
   	
%>
