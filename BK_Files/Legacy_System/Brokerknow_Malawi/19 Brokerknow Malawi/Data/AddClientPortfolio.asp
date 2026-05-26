<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add Client Portfolio </title>

	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
	<!--CALENDAR -->
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<script language="JavaScript" src="CALENDAR/calendar.js"></script>

	<!--END CALENDAR -->
	<script language="JavaScript" src="../Scripts/common.js"></script>
</head>

<body Class="Dialog">

<!--#include file="../libroutines.asp"-->
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<%
	
	Dim action
	Dim conn 
   Dim sqlStr
   Dim rs
	
	action = ucase(Request.Form("action"))
	
	if action = "EXECUTE" then
		Dim client
		Dim security
		Dim pDate
		Dim price
		Dim qty
        Dim Ref
        
       client = Request.Form("cboClient")
       security = Request.Form("cboSecurity")
		pDate = Request.Form("txtPDate")
       price = Request.Form("txtPrice")
       qty = Request.Form("txtQty")
       Ref=Request.Form("txtRef")
       
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
        sqlStr = "INSERT INTO [CPortfolio] (CPortfolioPDate,CPortfolioPrice,CPortfolioQty,CPortfolio_DPA_" & _
                ",Client_DPA_,Security_DPA_,Reference) SELECT " & "#" & FormatDate(PDate) & "#" & " as CPortfolioPDate" & _
                "," & " " & price & " " & " as CPortfolioPrice" & _
                "," & " " & qty & " " & " as CPortfolioQty," & " " & "iif(isnull(max([CPortfolio_DPA_])),(SELECT InitialID FROM _Initial_Table_ID_ WHERE TableName = 'CPortfolio'),max([CPortfolio_DPA_]) + 1)" & " " & " as CPortfolio_DPA_" & _
                "," & " " & client & " " & " as Client_DPA_" & _
                "," & " " & security & " " & " as Security_DPA_" & _
                "," & "'" & Ref & "'" & " as Reference" & _
                " FROM [CPortfolio]"
        Set conn = GetActiveConnection("KBroker")
        
        conn.BeginTrans
                conn.Execute SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
        conn.CommitTrans
        conn.Close
        Set conn = Nothing
        WritefraEnabledDialogCloseScript
        Response.End        
   	end If
%>
<form name = 'frmAddCPortfolio' method = 'post' action = 'AddClientPortfolio.asp' >
<table border="0" width="100%" cellspacing="1" cellpadding="1">
  <tr>
				<td width="10%">Client</td>
				<td><input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);">
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
					        Set conn = GetActiveConnection("KBroker")
					        
					        sqlStr = "SELECT * FROM FullClientList order by ClientName"
					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF
					                ClientName=rs.Fields("ClientName")
					                NameClient=Mid(ClientName,1,30)
					                %>					                        
					                        <option SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=NameClient%>" value = '<%=rs.Fields("Client_DPA_")%>'><%=NameClient%></option>

					                        <%rs.MoveNext
					                Loop
					        End If
					%>

					    </select>
				</td>
			</tr>
			<tr>
				<td width="10%">Security</td>
				<td><input type = 'text' name ='txtSecurityCode' id = 'txtSecurityCode' size="10" onBlur="txtval = this.value; selectItem(cboSecurity);">
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
					        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
					        If Not (rs.EOF Or rs.BOF) Then
					                rs.MoveFirst
					                Do Until rs.EOF
					                ClientName=rs.Fields("SecurityNameEx")
					                NameClient=Mid(ClientName,1,30)
					                %>					                        
					                        <option SearchCode = "<%=rs.Fields("SecurityCode")%>" SearchText = "<%=NameClient%>" value = '<%=rs.Fields("Security_DPA_")%>'><%=NameClient%></option>
					                        <%rs.MoveNext
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
			var cal=new ctlSpiffyCalendarBox("cal", "frmAddCPortfolio", "txtPDate","cmdDate","<%= FormatDate(Date) %>",1);
			cal.writeControl();
		</SCRIPT>
	</td>
  </tr>
  <tr>
    <td width="25%">Purchase Price</td>
    <td width="75%"><input type = 'text' name ='txtPrice' id = 'txtPrice' size="20"></td>
  </tr>
  <tr>
    <td width="25%">Quantity</td>
    <td width="75%"><input type = 'text' name ='txtQty' id = 'txtQty' size="20"></td>
  </tr>
  <tr>
    <td width="25%">Reference</td>
    <td width="75%"><input type = 'text' name ='txtRef' id = 'txtRef' size="20"></td>
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
		<input type = 'button' Class=Buttons name ='cmdClose' id = 'cmdClose' value=" Cancel " OnClick="JavaScript: window.parent.close();">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
	</td>
  </tr>
</table>
</form>
</body>

</html>
