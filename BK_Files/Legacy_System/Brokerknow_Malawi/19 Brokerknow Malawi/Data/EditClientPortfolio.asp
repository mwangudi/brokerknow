<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit Client Portfolio</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
	<SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
	<script language="JavaScript" src="CALENDAR/calendar.js"></script>	
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
</head>

<body Class="Dialog">
<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
   Dim ID
   Dim rsEdit
	
	action = ucase(Request.Form("action"))
	ID = Request("ID")

		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		window.self.close
                </script>
                <% response.end
        End If

	if action = "EXECUTE" then
		Dim client
		Dim security
		Dim pDate
		Dim price
		Dim qty
        
       client = Request.Form("cboClient")
       security = Request.Form("cboSecurity")
	   pDate = Request.Form("txtPDate")
       price = Request.Form("txtPrice")
       qty = Request.Form("txtQty")       
       toCancel = Request.Form("cmdCancel")
       Ref=Request.Form("txtRef")
       
       Set conn = GetActiveConnection("KBroker")
       
       If toCancel <> "" Then
			
			WriteDialogCancelScript
			Set Conn = Nothing
			Response.End
       End If
       
       
        'validate Client
        If Trim(Client) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Client"
                		
                </script>
                <% response.end
        End If
        'validate Security
        If Trim(Security) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Security"
                		
                </script>
                <% response.end
        End If
        'validate Purchase Price
        If Trim(Price) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Purchase Price"
                		
                </script>
                <% response.end
        End If
        'validate Quantity
        If Trim(qty) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "Please specify the Quantity"
                		
                </script>
                <% response.end
        End If
        'ensure Purchase Price is numeric
        If (Price <> "") And (Not IsNumeric(Price)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Purchase Price  must be numeric"
                
                </script>
                <% response.end
        End If
        
        'ensure Quantity is numeric
        If (Qty <> "") And (Not IsNumeric(Qty)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Quantity  must be numeric"
                
                </script>
                <% response.end
        End If

		
        'save data
        
        sqlStr = "UPDATE [CPortfolio] SET CPortfolioPDate = " & "#" & FormatDate(pDate)  & "#" & ",CPortfolioPrice = " & " " & price  & " " & "" & _
                ",CPortfolioQty = " & " " & qty & " " & ",Client_DPA_ = " & " " & client  & " " & ",Reference = " & "'" & Ref  & "'" & ",Security_DPA_ = " & " " & security & " " & "" & _
                " WHERE CPortfolio_DPA_  = " & ID
                
        
        conn.BeginTrans
                conn.Execute SQLServerFormat(HandleQuote(sqlStr))
        conn.CommitTrans
        
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End
   	end If
   	
%>

<form name = 'frmEditCPortfolio' method = 'post' action = 'EditClientPortfolio.asp' >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<% 
					Set conn = GetActiveConnection("KBroker")
					
					sqlStr = "SELECT CPortfolioPDate,CPortfolioPrice,CPortfolioQty,CPortfolio_DPA_" & _
					",ClientList.Client_DPA_,SecurityList.Security_DPA_,CPortfolio.Reference FROM [SecurityList] INNER JOIN ([ClientList] INNER JOIN [CPortfolio] ON ClientList.Client_DPA_ = CPortfolio.Client_DPA_) ON SecurityList.Security_DPA_ = CPortfolio.Security_DPA_ WHERE CPortfolio_DPA_  = " & ID
                
					Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					If rs.EOF Or rs.BOF Then%>
					        <script language = 'vbscript'>
					        		ShowMessage "The selected Client portfolio cannot be retrieved for editing"
					        		
					        </script>
					        <% response.end
					End If
					%>
  <tr>
				<td width="10%">Client</td>
				<td><input type = 'text' name ='txtClientCode' id = 'txtClientCode' value='<%=rs("Client_DPA_")%>' size="10" onBlur="txtval = this.value; selectItem(cboClient);">
				<select name = 'cboClient' id = "cboClient" size="1" 
    				onchange='UpdateCode(true,cboClient,txtClientCode)'
					onKeypress="return (dodefaultaction()==''); " 
					onKeydown="return (dodefaultaction()==''); " 
					onKeyup="return (FilterData(this,1,UpdateCode(change(cboClient,0),cboClient,txtClientCode)));" 
					onfocus="txtval = '';inputIsItemCode = 1;" 
					onblur="txtval = '';inputIsItemCode = 1;">
					<option selected SearchCode = "" SearchText = ""  value = ''></option>
					<%

					dim ClientName
					dim NameClient
					        
					        
					        sqlStr = "SELECT * FROM FullClientList order by ClientName"
					        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rsEdit.EOF Or rs.BOF) Then
					                rsEdit.MoveFirst
					                Do Until rsEdit.EOF
					                ClientName=rsEdit.Fields("ClientName")
					                NameClient=Mid(ClientName,1,30)
										if rsEdit.Fields("Client_DPA_") = rs.Fields("Client_DPA_") Then%>
										<option Selected SearchCode = "<%=rsEdit.Fields("Client_DPA_")%>" SearchText = "<%=NameClient%>" value = '<%=rsEdit.Fields("Client_DPA_")%>'><%=NameClient%></option>
					                
										<%
										else
										%>					                        
										        <option SearchCode = "<%=rsEdit.Fields("Client_DPA_")%>" SearchText = "<%=NameClient%>" value = '<%=rsEdit.Fields("Client_DPA_")%>'><%=NameClient%></option>

										        <%
										end if
					                rsEdit.MoveNext
					                Loop
					        End If
					%>

					    </select>
				</td>
			</tr>
			<tr>
				<td width="10%">Security</td>
				<td><input type = 'text' name ='txtSecurityCode' id = 'txtSecurityCode' value='<%=rs("Security_DPA_")%>' size="10" onBlur="txtval = this.value; selectItem(cboSecurity);">
				<select name = 'cboSecurity' id = "cboSecurity" size="1" 
    				onchange='UpdateCode(true,cboSecurity,txtSecurityCode)'
					onKeypress="return (dodefaultaction()==''); " 
					onKeydown="return (dodefaultaction()==''); " 
					onKeyup="return (FilterData(this,1,UpdateCode(change(cboSecurity,0),cboSecurity,txtSecurityCode)));" 
					onfocus="txtval = '';inputIsItemCode = 1;" 
					onblur="txtval = '';inputIsItemCode = 1;">
					<option selected SearchCode = "" SearchText = ""  value = ''></option>
					<%
					'dim ClientName
					'dim NameClient
					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT * FROM SecurityList order by SecurityName"
					        Set rsEdit = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rsEdit.EOF Or rsEdit.BOF) Then
					                rsEdit.MoveFirst
					                Do Until rsEdit.EOF
					                ClientName=rsEdit.Fields("SecurityNameEx")
					                NameClient=Mid(ClientName,1,30)
										if rsEdit.Fields("Security_DPA_") = rs.Fields("Security_DPA_") Then%>
										<option Selected SearchCode = "<%=rsEdit.Fields("SecurityCode")%>" SearchText = "<%=NameClient%>" value = '<%=rsEdit.Fields("Security_DPA_")%>'><%=NameClient%></option>					                					                        
										<%
										else
										%>
										<option SearchCode = "<%=rsEdit.Fields("SecurityCode")%>" SearchText = "<%=NameClient%>" value = '<%=rsEdit.Fields("Security_DPA_")%>'><%=NameClient%></option>
										        <%
										end if
					                rsEdit.MoveNext
					                Loop
					        End If
					%>

					    </select>
				</td>
			</tr>

   <tr>
    <td width="25%">Purchase Date</td>
    <td width="75%">
		<SCRIPT language="JavaScript">	
			var cal=new ctlSpiffyCalendarBox("cal", "frmEditCPortfolio", "txtPDate","cmdDate","<%= FormatDate(rs.Fields("CPortfolioPDate")) %>",1);
			cal.writeControl();
		</SCRIPT>
	</td>
  </tr>

  <tr>
    <td width="25%">Purchase Price</td>
    <td width="75%"><input type = 'text' name ='txtPrice' id = 'txtPrice' size="20" value = '<%=rs.Fields("CPortfolioPrice")%>'></td>
  </tr>
  <tr>
    <td width="25%">Quantity</td>
    <td width="75%"><input type = 'text' name ='txtQty' id = 'txtQty' value = '<%=rs.Fields("CPortfolioQty")%>' size="20"></td>
  </tr>
  <tr>
    <td width="25%">Reference</td>
    <td width="75%"><input type = 'text' name ='txtRef' id = 'txtRef' value = '<%=rs.Fields("Reference")%>' size="20"></td>
  </tr>
  
  <tr>
</table>
<table border=0 cellspacing=0 cellpadding=0 align=Right>
	<tr> 
    <td valign=absBottom>
		<BR>
		<BR>
		<BR>
		<BR>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdAdd' id = 'cmdAdd' value=" Save ">
		&nbsp;&nbsp;
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.self.close();">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
    	<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
	</td>
  </tr>
</table>  
</form>

</body>

