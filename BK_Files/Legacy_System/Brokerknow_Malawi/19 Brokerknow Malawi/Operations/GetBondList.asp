<!--BEGINPAGEFUNCS-->
<!--#include file="../libroutines.asp"-->

<%
const UDLName = "KBroker"

Set Conn = GetActiveConnection(UDLName)

whatToDo = Request.QueryString("Action")

Select Case LCase(Trim(whatToDo))
	Case "getaccountlist"
		Response.Write "<!--ENDPAGEFUNCS-->" &  GetAccountList			
		
	Case ""

End Select 

function GetAccountList
	Dim accList
	Dim rsAccount
	Dim currentEntityType
	Dim displayField
	
	currentEntityType = Request.QueryString("ID") 
	
	accList = "<option selected SearchCode = '' SearchText = ''  value = ''></option>"
	
	'if currentEntityType = 1 then 'client data
			displayField = "EntityName"
	'else
	'		displayField = "EntityNameEx"
	'end if
	
    sqlStr = "SELECT * FROM [CompleteEntityList] WHERE EntityType_DPA_ =" & currentEntityType & " Order By EntityName"
    Set rsAccount = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If Not (rsAccount.EOF Or rsAccount.BOF) Then
            rsAccount.MoveFirst
            Do Until rsAccount.EOF
                    accList = accList & "<option SearchCode = '" & rsAccount.Fields("EntityCode") & "' SearchText = '" & rsAccount.Fields("EntityName") & "'  value = '" & rsAccount.Fields("Entity_DPA_") & "'>" & rsAccount.Fields(displayField) & "</option>"
                    rsAccount.MoveNext
            Loop
    End If
    	
	GetAccountList = accList
end function
%>
