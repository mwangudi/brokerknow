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
	Case "sloadclient"
		Response.Write "<!--ENDPAGEFUNCS-->" &  SLoadClient
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


function LoadClient 
	Dim AgentReturnable
	Dim OrderContact
	Dim outputstr

	clientcode = Request.QueryString("clientcode")
	cdsno = Request.QueryString("cdsno")
	filterstr = ""

	if(trim(clientcode)<>"") then
		filterstr =" and client.client_DPA_=" & clientcode
	else
		if(trim(cdsno)<>"") then
			filterstr =" and client.clientCDSNo like '%" & cdsno & "%'"
		end if
	end if

	sqlStr = " SELECT     TOP 1 ISNULL(ClientBalances.CurrentBal, 0) + ISNULL(Client.CreditLimit, 0) - ISNULL(ClientTotal.Total, 0) AS AvailableCredit,  " & _
			"                       ISNULL(ClientBalances.CurrentBal, 0) AS CurrentBal, ClientTotal.Total, Agent.AgentName, OwnerList.OwnerName, ISNULL(Client.Agent_DPA_, 0)  " & _
			"                       AS Agent_DPA_, Client.Client_DPA_, Client.ClientContact, Client.ClientName, ISNULL(Client.Owner_DPA_, 0) AS Owner_DPA_, Client.ClientCDSNo,  " & _
			"                       Client.IsCustodian " & _
			" FROM         Client LEFT OUTER JOIN " & _
			"                       ClientTotal ON Client.Client_DPA_ = ClientTotal.Client_DPA_ LEFT OUTER JOIN " & _
			"                       ClientBalances ON Client.Client_DPA_ = ClientBalances.client_DPA_ LEFT OUTER JOIN " & _
			"                       OwnerList ON Client.Owner_DPA_ = OwnerList.Owner_DPA_ LEFT OUTER JOIN " & _
			"                       Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_ " & _
			" WHERE     (Client.Deleted = 0) AND (Client.IsFrozen <> 1) " & filterstr  
	 
	'response.write sqlStr
	'response.end

	set rs = conn.Execute(sqlStr)

	If Not (rs.EOF Or rs.BOF) Then
		OrderContact = rs.Fields("ClientContact").value
		Iscustodian = rs.Fields("IsCustodian").value
		AgentID = rs.Fields("Agent_DPA_").value
		Agent = rs.Fields("AgentName").value
		OwnerID = rs.Fields("Owner_DPA_").value
		Owner = rs.Fields("OwnerName").value
		Credit = FormatNum(rs.Fields("AvailableCredit").value)
	    CurrentBal = FormatNum(rs.Fields("CurrentBal").value)
		'Credit = rs.Fields("AvailableCredit").value
		'CurrentBal = rs.Fields("CurrentBal").value
		SearchCode = rs.Fields("Client_DPA_").value
		SearchText = Replace(rs.Fields("ClientName").value,";","")
		SearchCDS = rs.Fields("ClientCDSNo").value	  
	Else
		OrderContact = ""
		Iscustodian = 0
		AgentID = 0
		Agent = ""
		OwnerID = 0
		Owner = ""
		Credit = 0
		CurrentBal = 0
		SearchCode = 0
		SearchText = ""
		SearchCDS = ""
	End If

	outputstr = Credit & "<->" & CurrentBal & "<->" & Agent & "<->" & Owner & "<->" & AgentID & "<->" & SearchCode & "<->" & OrderContact & "<->" & SearchText & "<->" & OwnerID & "<->" & SearchCDS & "<->" & Cint(IsCustodian)
		 
	'response.write outputstr
	'response.end

	LoadClient = outputstr
end function

function SLoadClient
 
 Dim AgentReturnable
 Dim OrderContact
 Dim outputstr

 clientcode = Request.QueryString("clientcode")
 cdsno = Request.QueryString("cdsno")
 clientname = Request.QueryString("clientname")
 myguidstr = Request.QueryString("guidStr")
 
 filterstr = ""

 if clientname <> "" then
	filterstr =" and client.ClientName like '%" & clientname & "%'"
 end if

 if(trim(clientcode)<>"") then
 filterstr =" and client.client_DPA_=" & clientcode
 else
	if(trim(cdsno)<>"") then
	filterstr =" and client.clientCDSNo like '%" & cdsno & "%'"
	end if
 end if

	sqlStr =" SELECT  ISNULL(ClientBalances.CurrentBal, 0) + ISNULL(Client.CreditLimit, 0) - ISNULL(ClientTotal.Total, 0) AS AvailableCredit,  " & _
			"                       ISNULL(ClientBalances.CurrentBal, 0) AS CurrentBal, ClientTotal.Total, Agent.AgentName, OwnerList.OwnerName, ISNULL(Client.Agent_DPA_, 0)  " & _
			"                       AS Agent_DPA_, Client.Client_DPA_, Client.ClientContact, Client.ClientName, ISNULL(Client.Owner_DPA_, 0) AS Owner_DPA_, Client.ClientCDSNo,  " & _
			"                       Client.IsCustodian " & _
			" FROM Client LEFT OUTER JOIN " & _
			"                       ClientTotal ON Client.Client_DPA_ = ClientTotal.Client_DPA_ LEFT OUTER JOIN " & _
			"                       ClientBalances ON Client.Client_DPA_ = ClientBalances.client_DPA_ LEFT OUTER JOIN " & _
			"                       OwnerList ON Client.Owner_DPA_ = OwnerList.Owner_DPA_ LEFT OUTER JOIN " & _
			"                       Agent ON Client.Agent_DPA_ = Agent.Agent_DPA_ " & _
			" WHERE (Client.Deleted = 0)  " & filterstr  
				
'response.write sqlStr
' response.end

set rs = conn.Execute(sqlStr)

If clientname = "" then

	If Not (rs.EOF Or rs.BOF) Then
		OrderContact = rs.Fields("ClientContact").value
		Iscustodian = rs.Fields("IsCustodian").value
		AgentID = rs.Fields("Agent_DPA_").value
		Agent = rs.Fields("AgentName").value
		OwnerID = rs.Fields("Owner_DPA_").value
		Owner = rs.Fields("OwnerName").value
		Credit = FormatNum(rs.Fields("AvailableCredit").value)
		CurrentBal = FormatNum(rs.Fields("CurrentBal").value)
		'Credit = rs.Fields("AvailableCredit").value
		'CurrentBal = rs.Fields("CurrentBal").value
		SearchCode = rs.Fields("Client_DPA_").value
		SearchText = Replace(rs.Fields("ClientName").value,";","")
		SearchCDS = rs.Fields("ClientCDSNo").value	  
	Else
		OrderContact = ""
		Iscustodian = 0
		AgentID = 0
		Agent = ""
		OwnerID = 0
		Owner = ""
		Credit = 0
		CurrentBal = 0
		SearchCode = 0
		SearchText = ""
		SearchCDS = ""
	End If

	outputstr = Credit & "<->" & CurrentBal & "<->" & Agent & "<->" & Owner & "<->" & AgentID & "<->" & SearchCode & "<->" & OrderContact & "<->" & SearchText & "<->" & OwnerID & "<->" & SearchCDS & "<->" & Cint(IsCustodian) & "<->" & myguidstr

Else
	
	a_clidet = rs.GetRows
	a_size = (rs.RecordCount - 1)

	for i = 0 to a_size
		'outputstr = outputstr & a_clidet(0,i) & ";" & a_clidet(1,i) & ";" & a_clidet(2,i) & "|"
		outputstr = outputstr & a_clidet(0,i) & "<->" & a_clidet(1,i) & "<->" & a_clidet(3,i) & "<->" & a_clidet(4,i) & "<->" & a_clidet(5,i) & "<->" & a_clidet(6,i) & "<->" & a_clidet(7,i) & "<->" & a_clidet(8,i) & "<->" & a_clidet(9,i) & "<->" & a_clidet(10,i) & "<->" & Cint(a_clidet(11,i)) & "|"
	next
	
	outputstr = left(outputstr,len(outputstr) - 1)
	
End if

 'response.write outputstr
 'response.end

 SLoadClient = outputstr
end function



%>