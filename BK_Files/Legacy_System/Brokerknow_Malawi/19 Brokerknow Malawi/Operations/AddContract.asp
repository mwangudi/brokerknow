<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim orderLineRS
   Dim ID
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please select an order on which to base the contract"
                		
                </script>
                <% response.end
        End If
        
    Set conn = GetActiveConnection("KBroker")
    
    sqlStr = "SELECT * FROM [OrdDetailList] WHERE (OrdDetailList.Order_DPA_ = " & ID & ") AND (BalanceQty > 0)"
    Set orderLineRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If orderLineRS.EOF Or orderLineRS.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Order cannot be retrieved"
                		Dim urlPage
                		Dim url
                		Dim ref
                		
                		ref = CStr(document.referrer)
                		'urlPage = split(,"/")
                		'ShowMessage "knmkljn"'UBound(urlPage)
                		'urlPage(UBound(urlPage)) = "ContrctOrderList.asp"
                		'url = join(urlPage,"/")
                		'window.history.go url
                </script>
                <% response.end
    End If
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=orderLineRS.fields("OrdDetailType")%> Contract</title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<SCRIPT language="JavaScript">
	var calTDate=new ctlSpiffyCalendarBox("calTDate", "frmAddContract", "txtTDate","cmdTDate","<%=Date%>",1);
	var calDelDate=new ctlSpiffyCalendarBox("calDelDate", "frmAddContract", "txtDelDate","cmdDelDate","<%=Date%>",1);
	var calNCDate=new ctlSpiffyCalendarBox("calNCDate", "frmAddContract", "txtNCDate","cmdNCDate","<%=Date%>",1);
</SCRIPT>
<!--END CALENDAR -->
</head>

<body><!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<%
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim broker
		Dim odetail
		Dim status
		Dim aPayno
		Dim bPayNo
		Dim cPayNo
		Dim contractNo
		Dim cert
		Dim tDate
		Dim delDate
		Dim ncDate
		Dim ncert
		Dim price
		Dim qty
		Dim slip
        
       broker = Request.Form("cboBroker") 
       odetail = Request.Form("cboOrdDetail")
       status = Request.Form("cboStatus")
       aPayno = Request.Form("txtAPayNo")
       bPayNo = Request.Form("txtBPayNo")
		cPayNo = Request.Form("txtCPayNo")
		contractNo = Request.Form("txtCNumber")
       cert = Request.Form("txtName")
       tDate = Request.Form("txtTDate")
       delDate = Request.Form("txtDelDate")
       ncDate = Request.Form("txtNCDate")
       ncert = Request.Form("txtNCertificate")
       price = Request.Form("txtPrice")
       qty = Request.Form("txtQty")
       slip = Request.Form("txtSlipNo")


        'validate Broker
        If Trim(Broker) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Broker"
                
                </script>
                <% response.end
        End If
        'validate Order Detail
        If Trim(ODetail) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Order Detail"
                
                </script>
                <% response.end
        End If
        'validate Contract Status
        If Trim(Status) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Contract Status"
                
                </script>
                <% response.end
        End If
        'validate size of Agent Payment No
        If Len(APayNo) > 20 Then%>
                <script language = 'vbscript'>
                ShowMessage "Agent Payment No can only be 20 characters in length"
                
                </script>
                <% response.end
        End If
        'validate size of Broker Payment No
        If Len(BPayNo) > 20 Then%>
                <script language = 'vbscript'>
                ShowMessage "Broker Payment No can only be 20 characters in length"
                
                </script>
                <% response.end
        End If
        'validate Contract Number
        If Trim(contractNo) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Contract Number"
                
                </script>
                <% response.end
        End If
        'validate size of Contract Number
        If Len(contractNo) > 20 Then%>
                <script language = 'vbscript'>
                ShowMessage "Contract Number can only be 20 characters in length"
                
                </script>
                <% response.end
        End If
        'validate size of Client Payment No
        If Len(CPayNo) > 20 Then%>
                <script language = 'vbscript'>
                ShowMessage "Client Payment No can only be 20 characters in length"
                
                </script>
                <% response.end
        End If
        'validate size of Certificate
        If Len(Cert) > 20 Then%>
                <script language = 'vbscript'>
                ShowMessage "Certificate can only be 20 characters in length"
                
                </script>
                <% response.end
        End If
        'validate size of New Certificate
        If Len(NCert) > 20 Then%>
                <script language = 'vbscript'>
                ShowMessage "New Certificate can only be 20 characters in length"
                
                </script>
                <% response.end
        End If
        'validate Price
        If Trim(Price) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Price"
                
                </script>
                <% response.end
        End If
        'ensure Price is numeric
        If (Price<> "") And (Not IsNumeric(Price)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Price  must be numeric"
                
                </script>
                <% response.end
        End If
        'validate Quantity
        If Trim(Qty) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Quantity"
                
                </script>
                <% response.end
        End If
        'ensure Quantity is numeric
        If (Qty<> "") And (Not IsNumeric(Qty)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Quantity  must be numeric"
                
                </script>
                <% response.end
        End If
        'validate Slip Number
        If Trim(Slip) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Slip Number"
                
                </script>
                <% response.end
        End If
        'validate size of Slip Number
        If Len(Slip) > 20 Then%>
                <script language = 'vbscript'>
                ShowMessage "Slip Number can only be 20 characters in length"
                
                </script>
                <% response.end
        End If
        
        'calculate amounts
        Dim grossAmount  'this is the amount before application of Levies
        
        grossAmount = price * qty
        
        'save data
        Dim guidStr 
        Dim guid 
        set guid = server.CreateObject("NDUtils.CGUID")
        guidStr = guid.GenerateGUID
        sqlStr = "INSERT INTO [Contract] (ContractAPayNo,ContractBPayNo,ContractCNumber,ContractCPayNo" & _
        "               ,ContractCertificate,ContractDelDate,ContractNCDate,ContractNCertificate" & _
        "               ,ContractPrice,ContractQty,ContractSlipNo,ContractTDate" & _
        "               ,Contract_DPA_,Contract_EIT_,Broker_DPA_,OrdDetail_DPA_,Status_DPA_) SELECT " & "'" & APayNo & "'" & " as ContractAPayNo" & _
        "       ," & "'" & BPayNo & "'" & " as ContractBPayNo" & _
        "       ," & "'" & contractNo & "'" & " as ContractCNumber" & _
        "       ," & "'" & CPayNo & "'" & " as ContractCPayNo" & _
        "       ," & "'" & Cert & "'" & " as ContractCertificate" & _
        "       ," & "#" & FormatDate(DelDate) & "#" & " as ContractDelDate" & _
        "       ," & "#" & FormatDate(NCDate) & "#" & " as ContractNCDate" & _
        "       ," & "'" & NCert & "'" & " as ContractNCertificate" & _
        "       ," & " " & Price & " " & " as ContractPrice," & " " & Qty & " " & " as ContractQty" & _
        "       ," & "'" & Slip & "'" & " as ContractSlipNo" & _
        "       ," & "#" & FormatDate(TDate) & "#" & " as ContractTDate" & _
        "       ," & " " & "iif(isnull(max([Contract_DPA_])),1,max([Contract_DPA_]) + 1)" & " " & " as Contract_DPA_" & _
        "       ," & "'" & guidStr & "'" & " as Contract_EIT_," & " " & Broker & " " & " as Broker_DPA_" & _
        "       ," & " " & ODetail & " " & " as OrdDetail_DPA_" & _
        "       ," & " " & Status & " " & " as Status_DPA_" & _
        "        FROM [Contract]"
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
                'obtain contract key value
                Dim contractRS
                sqlStr = "SELECT [Contract.Contract_DPA_] FROM [Contract] WHERE [Contract.Contract_EIT_] = " & "'" & guidStr & "'"
                Set contractRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
                If (contractRS.EOF Or contractRS.BOF) Then%>
                <script language = 'vbscript'>
                        ShowMessage "A serious error has been encountered while saving the data. Try saving again"
                        
                </script>
                <% response.end
                End If
                'apply relevant levies
                Dim levyRS
                if orderLineRS.Fields("OrdDetailSecType")= "F" then
					cond = "WHERE LevyAppBond = True"
				else
					cond = "WHERE LevyAppSecurity = True"
                end if
                sqlStr = "SELECT * FROM [LevyList] " & cond
				Set levyRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
				
				Dim levyAmount
				If Not (levyRS.EOF Or levyRS.BOF) Then
						levyRS.MoveFirst
						Do Until levyRS.EOF
								if levyRS.Fields("LevyType") = "P" Then
										levyAmount = CCur((levyRS.Fields("LevyAmount")/100.00) * grossAmount)
								else
										Dim blocks
										blocks = Int(grossAmount / levyRS.Fields("LevyBlock"))
										if (grossAmount mod levyRS.Fields("LevyBlock")) <> 0 then
												blocks = blocks + 1
										end if
										levyAmount = CCur(blocks * levyRS.Fields("LevyAmount"))
								end if
								sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),1,max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
										"       ," & " " & contractRS.Fields("Contract.Contract_DPA_") & " " & " as Contract_DPA_" & _
										"       ," & "'" & levyRS.Fields("LevyDescription") & "'" & " as LevyName" & _
										"       ," & " " & levyAmount & " " & " as LevyAmount" & _
										"        FROM [LevyContract]"
								conn.Execute SQLServerFormat(HandleQuote(sqlStr))
								levyRS.MoveNext
						Loop
				end if
				'consider the transfer fee
				if Not(orderLineRS.Fields("OrderTypeSale")) then
					Dim transRS
					sqlStr = "SELECT * FROM [SecTransFeeListLatest] WHERE Security_DPA_ = " & orderLineRS.Fields("Security_DPA_")
					Set transRS = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),1,max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
							"       ," & " " & contractRS.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
							"       ," & "'Transfer Fee'" & " as LevyName" & _
							"       ," & " " & transRS.Fields("Fee") & " " & " as LevyAmount" & _
							"        FROM [LevyContract]"
					conn.Execute SQLServerFormat(HandleQuote(sqlStr))
				end if
				'Apply broker commission
				levyAmount = CCur((orderLineRS.Fields("CommissionRate")/100.00) * grossAmount)
				sqlStr = "INSERT INTO [LevyContract] (LevyContract_DPA_,Contract_DPA_,LevyName,LevyAmount) SELECT " & " " & "iif(isnull(max([LevyContract_DPA_])),1,max([LevyContract_DPA_]) + 1)" & " " & " as LevyContract_DPA_" & _
						"       ," & " " & contractRS.Fields("Contract_DPA_") & " " & " as Contract_DPA_" & _
						"       ," & "'Broker Commission'" & " as LevyName" & _
						"       ," & " " & levyAmount & " " & " as LevyAmount" & _
						"        FROM [LevyContract]"
				conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        response.redirect "ViewContract.asp?contractID=" & contractRS.Fields("Contract.Contract_DPA_") & "&contractType=new"
   	end If
%>
<CENTER>
	<DIV class="ListNugget" id="AdvSearchHead" style="WIDTH: 640px" name="AdvSearchHead">
		<TABLE class="ListNuggetHeader" cellPadding="0" cellSpacing="0" width="100%" name="AdvSearchtestHeader"> 
			<TR>
			<TD class="ListNuggetTitleCellWhite"
					onselectstart="window.event.cancelBubble=true; return false;"   
					onclick="PartWrapperToggle('AdvSearchHead');">
					<A class=ListNuggetTitle onclick="return PartWrapperToggle('AdvSearchHead');"  
					 href="javascript:PartWrapperToggle('AdvSearchHead');"><%=orderLineRS.fields("OrdDetailType")%> Contract
					</A>
				</TD>
			 
				<TD class=ListNuggetButtonCellWhite onclick="PartWrapperToggle('AdvSearchHead');">
				<DIV class=ListNuggetButton>
					<IMG class=ListNuggetUpButton id=AdvSearchUp height=17 alt="Hide options" src="../images/blue-chevron_up.gif" width=17 align=right border=0 name=AdvSearchHeadUp>
					<IMG class=ListNuggetDownButton id=AdvSearchDown height=17 alt=Options src="../images/gray-chevron_down.gif" width=17 align=right border=0 name=AdvSearchHeadDown>
				</DIV>
			</TD>
			</TR>
		</TABLE>
		
<DIV class="ListNuggetBody" id="AdvSearchHeadBody" name="AdvSearchHeadBody" style="WIDTH: 640px">
<table class="srch_bg" style="MARGIN-TOP: 0px" cellPadding="1" width=100% cellSpacing="0" border="0">  
<tr><td>
<form name = 'frmAddContract' method = 'post' action = 'AddContract.asp' >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20%">Counter Party</td>
    <td width="80%"><select name = 'cboBroker' id = 'cboBroker' size="1">
    	<option selected value = ''></option>
<%
      
        sqlStr = "SELECT * FROM [BrokerList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF%>
                        <option value = '<%=rs.Fields("Broker_DPA_")%>'><%=rs.Fields("BrokerName")%></option>
                        <%rs.MoveNext
                Loop
        End If
%>

    </select></td>
  </tr>
  <tr>
    <td width="20%">Order Reference</td>
    <td width="80%"><b></b><%=orderLineRS.Fields("OrderRef")%></b></td>
  </tr>
  <tr>
    <td width="20%">Order Line</td>
    <td width="80%"><select name = 'cboOrdDetail' id = 'cboOrdDetail' size="1">
    	<option selected value = ''></option>
<%
        If Not (orderLineRS.EOF Or orderLineRS.BOF) Then
                orderLineRS.MoveFirst
                Do Until orderLineRS.EOF%>
                        <option value = '<%=orderLineRS.Fields("OrdDetail_DPA_")%>'><%=orderLineRS.Fields("OrdDetailItem")%>	[Bal : <%=orderLineRS.Fields("BalanceQty")%>]</option>
                        <%orderLineRS.MoveNext
                Loop
        End If
%>

    </select></td>
  </tr>
  <tr>
    <td width="20%">Status</td>
    <td width="80%"><select name = 'cboStatus' id = 'cboStatus' size="1">
    	<option selected value = ''></option>
<%
        sqlStr = "SELECT * FROM [StatusList]"
        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
                Do Until rs.EOF%>
                        <option value = '<%=rs.Fields("Status_DPA_")%>'><%=rs.Fields("StatusDescription")%></option>
                        <%rs.MoveNext
                Loop
        End If
%>

    </select></td>
  </tr>
  <tr>
    <td width="20%">Contract Number</td>
    <td width="80%"><input type = 'text' name ='txtCNumber' id = 'txtCNumber' size="20"></td>
  </tr>
    <tr>
    <td width="20%">Price</td>
    <td width="80%"><input type = 'text' name ='txtPrice' id = 'txtPrice' size="20"></td>
  </tr>
  <tr>
    <td width="20%">Quantity</td>
    <td width="80%"><input type = 'text' name ='txtQty' id = 'txtQty' size="20"></td>
  </tr>
  <tr>
    <td width="20%">CDS Ref</td>
    <td width="80%"><input type = 'text' name ='txtSlipNo' id = 'txtSlipNo' size="20"></td>
  </tr>
  <tr>
    <td width="20%">Trade Date</td>
    <td width="80%"><SCRIPT language="JavaScript">calTDate.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td width="20%">Delivery Date&nbsp;</td>
    <td width="80%"><SCRIPT language="JavaScript">calDelDate.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td width="20%">Certificate No</td>
    <td width="80%"><input type = 'text' name ='txtCertificate' id = 'txtCertificate' size="20"></td>
  </tr>
  <tr>
    <td width="20%">Agent Payment No</td>
    <td width="80%"><input type = 'text' name ='txtAPayNo' id = 'txtAPayNo' size="20"></td>
  </tr>
  <tr>
    <td width="20%">Broker Payment No</td>
    <td width="80%"><input type = 'text' name ='txtBPayNo' id = 'txtBPayNo' size="20"></td>
  </tr>
  <tr>
    <td width="20%">Client payment No</td>
    <td width="80%"><input type = 'text' name ='txtCPayNo' id = 'txtCPayNo' size="20"></td>
  </tr>
  <tr>
    <td width="20%">New Certificate Date</td>
    <td width="80%"><SCRIPT language="JavaScript">calNCDate.writeControl();</SCRIPT></td>
  </tr>
  <tr>
    <td width="20%">New Certificate No</td>
    <td width="80%"><input type = 'text' name ='txtNCertificate' id = 'txtNCertificate' size="20"></td>
  </tr>	
  <tr>
    <td width="20%"><input type = 'submit' class=buttons name ='cmdAdd' id = 'cmdAdd' value="Save"></td>
    <td width="80%"><input type = 'hidden' name ='action' id = 'action' value="Execute"></td>
    <td width="80%"><input type = 'hidden' name ='ID' id = "ID" value="<%=ID%>"></td>
  </tr>
</table>
</form>
</td>
</tr>
</table>

</div>
</div>

</body>

</html>








