<!--#include file="../libroutines.asp"-->
<%
''DB Connection
Dim Conn 
    Set Conn = CreateObject("ADODB.Connection")
    theDBName = "KBroker" 
    Conn.ConnectionString =  "FILE NAME=" & GetUDLPath(theDBName) 
    Conn.Open
    
   Function GetUDLPath(theDBName) 
    Dim tmpStr
    
    tmpStr = StrReverse(Request.ServerVariables("APPL_PHYSICAL_PATH"))
    
    tmpStr = Mid(tmpStr, InStr(1, tmpStr, "\") + 1)
    
    tmpStr = StrReverse(tmpStr)
    
    GetUDLPath = tmpStr & "\UDL\" & Trim(theDBName) & ".UDL"

End Function

Dim action
 
'Bond Parameters

Dim Client
Dim BondType
Dim BondIssue
Dim TradeType
Dim TradeDate
Dim Price
Dim Quantity
Dim Reference


action= ucase(Request.Form("action"))


if action = "CALCULATE" then
   Client = trim(Request.Form("cboClient"))
   BondType = trim(Request.Form("BondType"))
   BondIssue = trim(Request.Form("BondIssue"))
   TradeType = Cdate(trim(Request.Form("TradeType")))
   TradeDate = Cdate(trim(Request.Form("TradeDate")))
   Price = trim(Request.Form("Price"))
   Quantity = trim(Request.Form("Quantity"))
   Reference = trim(Request.Form("Reference"))
 end if

%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>BOND CALCULATOR</title>
 
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>

<!--CALENDAR -->

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
 <LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css"> 
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT>
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->


<script language='javascript'>
	
		var validNavigate = false;
		function ReleaseRecord()
		{
			if(!validNavigate)
			{
 				event.returnValue = "Please use the cancel button to close the dialog"
 			}
		}
		
		function AllowedNavigation()
		{
			validNavigate = true;
		}
		
				function chkDate(){
			var item=document.frmBondTrades 
			if(item.ValidityCheck.checked==true){
				item.ValidityDate.value=item.ValidDate.value;
			}
			else
			{
				item.ValidityDate.value='';
			}
		}
		function Format2Decimals(theItem)
	{
		var item = document.frmBondTrades;
			theItem.value=format2NumberCommasOnly(theItem.value)
			theItem.value= formatNum(theItem.value);
	 //document.all.item("txtTotal").value= formatNum(document.all.item("txtTotal").value);
	}
	function formatnumber(theTxt)
	{	
	var theprice = theTxt.value;
		theTxt.value=formatNum(theprice);
	}
	
function format_number(p,d) 
	{
  	var r;
  	if(p<0)
  		{
  		p=-p;
  		r=format_number2(p,d);
  		r="-"+r;
  		}
  	else
  		{
  		r=format_number2(p,d);
  		}
  return r;
	}

function format_number2(pnumber,decimals) 
	{
  	var strNumber = new String(pnumber);
  	var arrParts = strNumber.split('.');
  	var intWholePart = parseInt(arrParts[0],10);
  	var strResult = '';
  	if (isNaN(intWholePart))
    intWholePart = '0';
  	if(arrParts.length > 1)
  		{
    	var decDecimalPart = new String(arrParts[1]);
    	var i = 0;
    	var intZeroCount = 0;
     	while ( i < String(arrParts[1]).length )
     		{
       		if( parseInt(String(arrParts[1]).charAt(i),10) == 0 )
       			{
         		intZeroCount += 1;
         		i += 1;
       			}
       		else
         	break;
    		}
    	decDecimalPart = parseInt(decDecimalPart,10)/Math.pow(10,parseInt(decDecimalPart.length-decimals-1)); 
    	Math.round(decDecimalPart); 
    	decDecimalPart = parseInt(decDecimalPart)/10; 
    	decDecimalPart = Math.round(decDecimalPart); 

    	//If the number was rounded up from 9 to 10, and it was for 1 'decimal' 
    	//then we need to add 1 to the 'intWholePart' and set the decDecimalPart to 0. 

    	if(decDecimalPart==Math.pow(10, parseInt(decimals)))
    		{ 
      		intWholePart+=1; 
      		decDecimalPart="0"; 
    		} 
    	var stringOfZeros = new String('');
    	i=0;
    	if( decDecimalPart > 0 )
    		{
      		while( i < intZeroCount)
      			{
        		stringOfZeros += '0';
        		i += 1;
      			}
    		}
    	decDecimalPart = String(intWholePart) + "." + stringOfZeros + String(decDecimalPart); 
    	var dot = decDecimalPart.indexOf('.');
    	if(dot == -1)
    		{
      		decDecimalPart += '.'; 
      		dot = decDecimalPart.indexOf('.'); 
    		} 
    	var l=parseInt(dot)+parseInt(decimals); 
    	while(decDecimalPart.length <= l) 
    		{
      		decDecimalPart += '0'; 
    		}
    	strResult = decDecimalPart;
  		}
  	else
  		{
    	var dot; 
    	var decDecimalPart = new String(intWholePart); 

    	decDecimalPart += '.'; 
    	dot = decDecimalPart.indexOf('.'); 
    	var l=parseInt(dot)+parseInt(decimals); 
    	while(decDecimalPart.length <= l) 
    		{
      		decDecimalPart += '0'; 
    		}
    	strResult = decDecimalPart;
  		}
  	return strResult;
	}

	
</script>
<script language="Javascript">
		function validate(){
			var item=document.frmBondTrades
			
			
			if(item.txtClientCode.value==''){
				alert('Please select a client');
				item.txtClientCode.focus;				
				return false;			
			}
			if(item.TradeType.selectedIndex==''){
				alert('Please select a Trade Type');
				item.TradeType.focus;				
				return false;			
			}
			if(item.BondType.selectedIndex==''){
				alert('Please select a Bond Type');
				item.BondType.focus;				
				return false;			
			}
			if(item.IssueNo.selectedIndex==''){
				alert('Please select a Bond Issue');
				item.IssueNo.focus;				
				return false;			
			}
			
			if(item.TradeDate.value==''){
				alert('Please select the trade date');
				item.TradeDate.focus;				
				return false;			
			}

			if(item.Price.value==''){
						alert('Please input the price');
						item.Price.focus;				
						return false;			
					}
			if(item.Quantity.value==''){
						alert('Please input the Quantity');
						item.Quantity.focus;				
						return false;			
					}
			if(item.Reference.value==''){
						alert('Please input the Reference');
						item.Reference.focus;				
						return false;			
					}
			return true;
		}
		
function FetchAccounts(theList)
		{
			var i = 0;
			var entity = theList.value;
			var toList = document.frmBondTrades.IssueNo;
			var issueno = document.frmBondTrades.HidIssueNo.value;								
			
			frm = document.frmBondTrades;				
			xmlhttp = createXMLHTTPObj();
			
			url="GetList.asp?ID="+entity+"&action=GetBondList&issueno="+issueno;
			xmlhttp.open("GET",url,true);
			xmlhttp.onreadystatechange=function() {
				if (xmlhttp.readyState==4) {
				returnStr = xmlhttp.responseText;
				returnStr = getBodyHTML(returnStr);
			
				var secList = "<select name = '" + toList.name + "' id = '" + toList.name + "' size='1'' ";
				secList += "OnClick='event.cancelBubble=true;' " ;
				secList += "onChange='event.cancelBubble=true;' " ;
				secList += "onKeypress='return (dodefaultaction()==\"\"); ' "  ;
				secList += "onKeydown='return (dodefaultaction()==\"\");' " ; 
				secList += "onKeyup='return (change(" + toList.name + "));' " ; 
				secList += "onfocus='txtval = \"\";inputIsItemCode = 1;' "  ;
				secList += "onblur='txtval = \"\";inputIsItemCode = 1;'>" ;
				secList += returnStr ;
				secList += "</select>";
				
				toList.outerHTML = secList;														
				}
				}
			xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
			xmlhttp.send(); 
		
	}
	
	
</script>
</head>
<body leftmargin="20"  Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<%
'Defalt Dates
  
 if isnull(TradeDate) or trim(TradeDate)="" then TDate = "" else TDate = formatdate(TradeDate)
 
%>
<SCRIPT language="JavaScript">

	var cal=new ctlSpiffyCalendarBox("cal", "frmBondTrades", "TradeDate","cmdTradeDate","<%=TDate %>",1);
</SCRIPT>
<form name="frmBondTrades" method = 'post' action = 'AddBondTradeDetail.asp' id = "frmBondTrades" onsubmit="return validate();">
<p><b>BOND TRADES</b></p>



<table border="1" width="90%" cellspacing="0" cellpadding="0" bordercolor="#000000">
  <tr>
    <td>&nbsp;Client</td>
	<td>
	   <input type = 'text' name ='txtClientCode' id = 'txtClientCode' size="10" onBlur="txtval = this.value; selectItem(cboClient);" value="<%=client%>">
	<select name = 'cboClient' id = 'cboClient' size="1" 
			onKeypress="return (dodefaultaction()==''); " 
			onKeydown="return (dodefaultaction()==''); " 
			onKeyup="return (UpdateCode(change(cboClient,0),cboClient,txtClientCode));" 
			onChange="UpdateCode(true,cboClient,txtClientCode);"
			onfocus="txtval = '';inputIsItemCode = 1;" 
			onblur="txtval = '';inputIsItemCode = 1;" readonly >
		<!--option selected SearchCode = "" SearchText = "" value = ''></option-->
		<option selected SearchCode = "" SearchText = "" value = '' AccManager=""></option>
		<%
		set rs =server.createObject("Adodb.recordset")
		dim ClientName
		dim NameClient
		set rs = server.CreateObject("Adodb.recordset")
	   
		
		sqlStr =" SELECT * from fullclientlist"

		Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		If Not (rs.EOF Or rs.BOF) Then
				rs.MoveFirst
				Do Until rs.EOF					                
				
				ClientName=rs.Fields("ClientName")
				''NameClient=rs.Fields("Client_DPA_") & " " & Mid(ClientName,1,20)
						%>                    
						<option  SearchCode = "<%=rs.Fields("Client_DPA_")%>" SearchText = "<%=rs.Fields("ClientName")%>" value = '<%=rs.Fields("Client_DPA_")%>' 
						<%if (rs.Fields("Client_DPA_")=clng(client))then
							Response.Write "selected"
						 end if%> ><%=mid(ClientName,1,30)%></option>
						<%rs.MoveNext
				Loop
		End If
			  
		%>
	</select></td>
  </tr>
    
  <tr>
    <td width="30%">&nbsp;Bond Type</td>
    <td width="70%">
    <select name="BondType" id="BondType" size ='1' onchange='FetchAccounts(this);'>
    
    <%
		dim rst
		set rst = server.CreateObject("Adodb.recordset")
	    sqlStr = "SELECT * FROM [Security] where OrderSecType_DPA_ = 1 order by Security_DPA_"
	
        Set rst = Conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
		
    %>
		<option></option>
		<% while not rst.eof %>
		<option  value="<%=rst("Security_DPA_")%>" <%if(rst("Security_DPA_")=clng(BondType)) then Response.Write "Selected"%>><%=rst("SecurityCode")%></option>
	<%
		rst.movenext
		wend
		rst.close
		set rst= nothing	
	%>
	</select></td>
  </tr>

   <tr>
    <td>&nbsp;Bond Issue</td>
    <td>
    <select name = 'IssueNo' id = 'IssueNo' size="1" onchange="IssueOnChange();">
		<option selected SecCode = '0' SearchCode = "0" SearchText = ""  value = '' cboIDate='' cboMDate=''
		cboCPayments='' cboRate='' cboFaceValue=''   >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
	<%
		if BondType="" then BondType=0
		sqlStr = "SELECT  IssueList.* From [IssueList] WHERE Security_DPA_ =" & BondType         
    	dim rsAccount
		set rsAccount= server.CreateObject("Adodb.recordset")
    	Set rsAccount = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If Not (rsAccount.EOF Or rsAccount.BOF) Then
            rsAccount.MoveFirst
            Do Until rsAccount.EOF
                                   
               %> <option 		SecCode = '<%= rsAccount.Fields("Bond_DPA_")%>'
								SearchText = '<% = rsAccount.Fields("BondIssue")%>'
								value = '<% = rsAccount.Fields("BondIssue")%>'
								cboIDate='<% =  rsAccount.Fields("BondIDate")%>'
								cboMDate='<% =  rsAccount.Fields("BondMDate") %>'
								cboCPayments='<% = rsAccount.Fields("BondPayment")%>'
								cboRate='<% =  rsAccount.Fields("BondRate") %>'
								cboFaceValue='<% =  formatnumber(rsAccount.Fields("FaceValue"),2)%>'
								<% if(rsAccount.Fields("Bond_DPA_")=cint(SecurityCode)) then response.write "selected"%>
								>
								<%= rsAccount.Fields("BondIssue")%> </option>
                   <% rsAccount.MoveNext
            Loop
    else
		
   
    End If
    	
	%>
    </select></td>
  </tr>
    <tr>
    <td width="30%">&nbsp;Trade Type</td>
    <td width="70%"><select name="TradeType">
		<option></option>
		<%
		set rs = Conn.Execute("select * from OrderType") 
		if TradeType="" then TradeType=2
		while not rs.eof 
		
		%>
		
		<option value="<%=rs("OrderType_DPA_")%>" <%if(rs("OrderType_DPA_")=cint(TradeType)) then Response.Write "selected"%>><%=rs("OrderTypeDescription")%></option>
		<%
			rs.movenext
			wend
		%>
	</select></td>
  </tr>
   <tr>
    <td width="30%">&nbsp;Trade Date</td>
    <td width="70%"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT>&nbsp;</td>
   
   
  <tr>
    <td width="30%">&nbsp;Price</td>
    <td width="70%"><input type="text" name="Price" size="39" Value="<%=Price%>"></td>
  </tr>
   <tr>
    <td width="30%">&nbsp;Quantity</td>
    <td width="70%"><input type="text" name="Quantity" size="39" Value="<%=Quantity%>"></td>
  </tr>	
 
   <tr>
    <td width="30%">&nbsp;Reference</td>
    <td width="70%">
    <input type="text" name="Reference" size="39" Value="<%=Reference%>"></td>
  </tr>	
</table>
<p><input type="Submit" value="SAVE" name="buttonAction">
<input type="Submit" value="CLOSE" name="buttonAction" onclick="javascript: self.close();">
<input type="hidden" value="<%=InputID%>" name="InputID">
<input type="hidden" value="<%=FormatDate(now)%>" name="ValidDate">
<input type="hidden" value="CALCULATE" name="action">
<input type="Hidden" name="HidIssueNo" id="HidIssueNo" size="1" Value="<%=IssueNo%>"></p>
</form>

</body>

</html>