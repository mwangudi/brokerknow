<!--BEGINPAGEFUNCS-->
<!--#include file="../libroutines.asp"-->

<%
const UDLName = "KBroker"

Set Conn = GetActiveConnection(UDLName)

whatToDo = Request.QueryString("Action")

Select Case LCase(Trim(whatToDo))
	Case "getaccountlist"
		Response.Write "<!--ENDPAGEFUNCS-->" &  GetAccountList			
	Case "getbondlist"
		Response.Write "<!--ENDPAGEFUNCS-->" &  GetBondList					
	Case ""

End Select 

function GetAccountList
	Dim accList
	Dim rsAccount
	Dim currentEntityType
	Dim displayField
	
	currentEntityType = Request.QueryString("ID") 
	
	accList = "<option selected SearchCode = '' SearchText = ''  value = ''></option>"
	
	DisplayName=""
	
    sqlStr = "SELECT * FROM [CompleteEntityList] WHERE EntityType_DPA_ =" & currentEntityType & " Order By EntityName"
    Set rsAccount = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If Not (rsAccount.EOF Or rsAccount.BOF) Then
            rsAccount.MoveFirst
            Do Until rsAccount.EOF 
            		AccountName=Mid(rsAccount("EntityName"),1,30)
			  if(rsAccount("EntityType_DPA_")=5) then
			  accList = accList & "<option SearchCode = '" & rsAccount.Fields("EntityCode") & "' SearchText = '" & AccountName & "'  value = '" & rsAccount.Fields("Entity_DPA_") & "'>" & AccountName & "</option>"			                       
			  else           
                    accList = accList & "<option SearchCode = '" & rsAccount.Fields("Entity_DPA_") & "' SearchText = '" & AccountName & "'  value = '" & rsAccount.Fields("Entity_DPA_") & "'>" & AccountName & "</option>"
			  end if
                    rsAccount.MoveNext
            Loop
    End If   
		
	GetAccountList = accList
end function
function GetBondList	
	Dim accList
	Dim rsAccount
	Dim currentEntityType
	Dim displayField
	
	currentEntityType = Request.QueryString("ID") 
	
	'accList = "<option selected SearchCode = '' SearchText = ''  value = ''></option>"
	accList=""	
			displayField = "Security"
	
    sqlStr = "SELECT DISTINCT BondIssue AS Issue,Security_DPA_,SecurityCode,Bond_DPA_ From [IssueList] WHERE Security_DPA_ =" & currentEntityType & " Order By SecurityCode"        
    
    Set rsAccount = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If Not (rsAccount.EOF Or rsAccount.BOF) Then
            rsAccount.MoveFirst
            Do Until rsAccount.EOF
                    accList = accList & "<option SearchCode = '" & rsAccount.Fields("SecurityCode") & "' SearchText = '" & rsAccount.Fields("Issue") & "'  value = '" & rsAccount.Fields("Bond_DPA_") & "'>" & rsAccount.Fields("Issue") & "</option>"
                    rsAccount.MoveNext
            Loop
    else
    accList = accList & "<option SearchCode = '' SearchText = ''  value = ''></option>"
    End If
    	
	GetBondList = accList
end function

%>