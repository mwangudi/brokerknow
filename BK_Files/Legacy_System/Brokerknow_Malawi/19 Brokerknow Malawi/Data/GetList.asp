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
	Case "getbondissues"
    	Response.Write "<!--ENDPAGEFUNCS-->" &  GetBondIssues					
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
            		AccountName=Mid(rsAccount("EntityNameEx"),1,30)           
                    accList = accList & "<option SearchCode = '" & rsAccount.Fields("Entity_DPA_") & "' SearchText = '" & AccountName & "'  value = '" & rsAccount.Fields("Entity_DPA_") & "'>" & AccountName & "</option>"
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
	
	accList=""	
	accList = accList & "<option selected SecCode='0' SearchCode = '0' SearchText = '' value = '' cboIDate='' cboMDate='' cboCPayments='' cboRate='' cboFaceValue=''></option>"
	
			displayField = "Security"
	if trim(currentEntityType)="" then currentEntityType=0
	
    sqlStr = "SELECT  * From [IssueList] WHERE Security_DPA_ =" & currentEntityType & "  and  bondMdate> getdate()Order By SecurityCode"        

    Set rsAccount = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If Not (rsAccount.EOF Or rsAccount.BOF) Then
            rsAccount.MoveFirst
            Do Until rsAccount.EOF
                    'accList = accList & "<option SearchCode = '" & rsAccount.Fields("SecurityCode") & "' SearchText = '" & rsAccount.Fields("Issue") & "'  value = '" & rsAccount.Fields("Issue") & "'>" & rsAccount.Fields("Issue") & "</option>"
                    
                    accList = accList & "<option SearchCode = '" & rsAccount.Fields("SecurityCode") & "' "&_
										" SecCode = '" & rsAccount.Fields("Bond_DPA_") & "'" &_
										" SearchText = '" & rsAccount.Fields("BondIssue") & "'" &_
										" value = '" & rsAccount.Fields("BondIssue") & "'" &_
										" cboIDate='"& FormatDate(rsAccount.Fields("BondIDate")) &"'" &_
										" cboMDate='"& FormatDate(rsAccount.Fields("BondMDate")) &"'" &_
										" Bond_DPA='"& FormatDate(rsAccount.Fields("Bond_DPA_")) &"'" &_
										" cboCPayments='"& rsAccount.Fields("BondPayment") &"'" &_
										" cboRate='"& rsAccount.Fields("Rate") &"'" &_
										" cboFaceValue='"& formatnumber(rsAccount.Fields("FaceValue"),2) &"'>" &_
										"" & rsAccount.Fields("BondIssue") & "</option>"
                    rsAccount.MoveNext
            Loop
    else
    accList = accList & "<option selected SecCode ='0' SearchCode = '0' SearchText = '' value = '' cboIDate='' cboMDate='' cboCPayments='' cboRate='' cboFaceValue=''>None</option>"
   
    End If
    	
	GetBondList = accList
end function

Function GetBondIssues	
	Dim issueList
	Dim rsIssue
	Dim currentBondType
	
	set rsIssue = server.CreateObject("ADODB.Recordset")
	currentBondType = Request.QueryString("ID") 
	
	issueList=""	
	issueList = issueList & "<option selected value = ''></option>"
	
    sqlStr = "SELECT  * From [IssueList] WHERE Security_DPA_ =" & currentBondType & " Order By BondIssue"        
    	
    Set rsIssue = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If Not (rsIssue.EOF Or rsIssue.BOF) Then
            rsIssue.MoveFirst
            Do Until rsIssue.EOF
                    
                    issueList = issueList & "<option value = '" & rsIssue.Fields("Bond_DPA_") & "'> " &_
										"" & rsIssue.Fields("BondIssue") & "</option>"
                   rsIssue.MoveNext
            Loop
            rsIssue.Close
            set rsIssue = nothing
    End If
    
	GetBondIssues = issueList
End function

'getting the account manager
function GetAccManager
	Dim accList
	Dim rsAccount
	Dim currentEntityType
	Dim displayField
	
	currentEntityType = Request.QueryString("ID") 
	
	accList = "<option selected SearchCode = '' SearchText = ''  value = ''></option>"
	
	DisplayName=""
	
    sqlStr = "SELECT * FROM [OwnerList] WHERE Client_DPA_ =" & currentEntityType & " Order By OwnerName"
    Set rsAccount = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If Not (rsAccount.EOF Or rsAccount.BOF) Then
            rsAccount.MoveFirst
            Do Until rsAccount.EOF 
					AccountName=Mid(rsAccount("OwnerName"),1,30)           
					if (rsAccount.Fields("Owner_DPA_") =clng(AccManager)) Then
	                    accList = accList & "<option selected SearchCode = '" & rsAccount.Fields("Owner_DPA_") & "' SearchText = '" & rsAccount.Fields("OwnerName") & "'  value = '" & rsAccount.Fields("Owner_DPA_") & "'>" & rsAccount.Fields("OwnerName") & "</option>"
	                 else 
						accList = accList & "<option  SearchCode = '" & rsAccount.Fields("Owner_DPA_") & "' SearchText = '" & rsAccount.Fields("OwnerName") & "'  value = '" & rsAccount.Fields("Owner_DPA_") & "'>" & rsAccount.Fields("OwnerName") & "</option>"
					end if
                    rsAccount.MoveNext
            Loop
    End If   
    
	GetAccManager = accList
end function

%>