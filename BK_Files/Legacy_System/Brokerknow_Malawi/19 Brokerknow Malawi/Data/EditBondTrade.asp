<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "BondsTradeList"
	const DataEntity = "BondsTrade"
	const DataEntityPlural = "BondsTrades"
	const ActionFolder = "Data"
	
	Dim UserId
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim rst
	Dim guid
	Dim guidStr
	Dim ID
	
	set rst = server.CreateObject ("ADODB.Recordset")
	
	UserId=Session("UserID")
	action = ucase(Request.Form("action"))
	ID = Request("ID")

    If Trim(ID) = "" Then%>
       <script language = 'vbscript'>
         ShowMessage "No record specified for editing"	
       </script>
                <% response.end
    End If
            
	if action = "EXECUTE" then
            
		    Dim client
			Dim BondType
			Dim BondIssue
			Dim TradeType
		    Dim TradeDate
			Dim Price
			Dim Quantity
			Dim Reference
			    
			code = Request.Form("code")
			client = Request.Form("client")
			BondType = Request.Form("cboBondType")
			BondIssue = Request.Form("cboBondIssue")
			TradeType = Request.Form("TradeType")
			TradeDate = Request.Form("txtDate")
			Price = Request.Form("Price")
			Quantity = Request.Form("Quantity")
			Reference = Request.Form("Ref")
			
			if  code = "" then
			 %><script language = 'Javascript'>
					alert('Please specify the Client Code.')
					window.history.back()
				</script><% response.end
			end if     
			
			if  client = "" or client = 0 then
			 %><script language = 'Javascript'>
					alert('Please specify the Client.')
					window.history.back()
				</script><% response.end
			end if 
			
			if  BondType = "" then
			 %><script language = 'Javascript'>
					alert('Please specify the Bond Type.')
					window.history.back()
				</script><% response.end
			end if 
			
			if  TradeType = "" then
			 %><script language = 'Javascript'>
					alert('Please specify the Trade Type.')
					window.history.back()
				</script><% response.end
			end if 
			
			if  TradeDate = "" then
			 %><script language = 'Javascript'>
					alert('Please specify the Trade Date.')
					window.history.back()
				</script><% response.end
			end if 
			
			if  Price = "" then
			 %><script language = 'Javascript'>
					alert('Please specify the Price.')
					window.history.back()
				</script><% response.end
			end if 
			
			if not  isnumeric(Price) then
			 %><script language = 'Javascript'>
					alert('Invalid Price.')
					window.history.back()
				</script><% response.end
			end if
			
			if  Quantity = "" then
			 %><script language = 'Javascript'>
					alert('Please specify the Quantity.')
					window.history.back()
				</script><% response.end
			end if 
			
			if not  isnumeric(Quantity) then
			 %><script language = 'Javascript'>
					alert('Invalid Quantity.')
					window.history.back()
				</script><% response.end
			end if
			
			
			'save data
					
		     Set conn = GetActiveConnection("KBroker")
				 conn.BeginTrans		
												      
				    sqlStr = "Update [BondTrades] set Client_DPA_ = " & client & ", Bond_DPA_ =" & BondIssue & ", " & _
									"OrderType_DPA_ = " & TradeType & ", TradeDate = #" & FormatDate(TradeDate) & "#, " & _
									"Price = " & Ccur(Price) & ", Quantity = " & Ccur(Quantity) & " , Reference = '" & Reference & "', " &_
									"ModifiedBy = " & session("UserID") & " ,DateModified = #" & FormatDate(now()) & "# Where BondTrades_DPA_=" & ID
									
					sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))

					conn.Execute sqlStr
					conn.CommitTrans
					
					WritefraEnabledDialogCloseScript
		
				    conn.Close
					Set conn = Nothing
					
					Response.End	
   	End if
   	
 Set conn = GetActiveConnection("KBroker")
	sqlStr = "SELECT * FROM BondsTradeList WHERE BondTrades_DPA_=" & ID
			        
	Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
    If rs.EOF Or rs.BOF Then%>
	   <script language = 'vbscript'>
	    window.self.ShowMessage "The selected <%=DataEntity%> cannot be retrieved for editing"             		
	   </script>
		<% response.end
	 End If
	   	
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Edit <%=DataEntity%></title>

<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
 <SCRIPT language=Javascript src="../scripts/accountList.js"></SCRIPT>
 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
<!--CALENDAR -->
<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
<script language="JavaScript" src="CALENDAR/calendar.js"></script>
<!--END CALENDAR -->

<script language='vbscript'>

					function EntitySelected(itemID)
 							frm<%=DataSource%>.elements("ID").value = itemID
 							frm<%=DataSource%>.elements("action").value = "Fetch_Accounts"
 							frm<%=DataSource%>.submit
 							
					end function
		
</script>
<script language='javascript'>
		var validNavigate = false;
		
		
		function UpdateClientCode(){
		 var item = document.frm<%=DataSource%>
		 item.code.value = item.client[item.client.selectedIndex].value;
		 
		}
		
		function FetchIssues(theList)
		{
			var entity = theList.value;
			var toList = document.frm<%=DataSource%>.cboBondIssue;								
						
			xmlhttp = createXMLHTTPObj();		
		    url="GetList.asp?ID="+entity+"&action=GetBondIssues";
			xmlhttp.open("GET",url,true);
			xmlhttp.onreadystatechange=function() {
				if (xmlhttp.readyState==4) {
				returnStr = xmlhttp.responseText;
				returnStr = getBodyHTML(returnStr);
			
				var secList = "<select name = '" + toList.name + "' id = '" + toList.name + "' size='1'>";
				secList += returnStr ;
				secList += "</select>";
				
				toList.outerHTML = secList;														
				}
				}
			xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
			xmlhttp.send(); 
	}
	
	function Validate(){
	  var item = document.frm<%=DataSource%>
	   
	  if(item.code.value==''){
				alert('Please specify the Client Code');
				item.code.focus;
				return false;			
		}
		
	  if(item.client.selectedIndex==''){
				alert('Please specify the Client');
				item.client.focus;
				return false;			
		}
		
	 if(item.cboBondIssue.selectedIndex==''){
				alert('Please specify the Bond Issue');
				item.cboBondIssue.focus;
				return false;			
		}
	
	if(item.cboBondType.selectedIndex==''){
				alert('Please specify the Bond Type');
				item.cboBondType.focus;
				return false;			
		}
	
	if(item.TradeType.selectedIndex==''){
				alert('Please specify the Trade Type');
				item.TradeType.focus;
				return false;			
		}
		
	if(item.txtDate.value==''){
				alert('Please specify the Trade Date');
				item.txtDate.focus;
				return false;			
		}
	if(item.Price.value==''){
				alert('Please specify the Price');
				item.Price.focus;
				return false;			
		}
		
    if(item.Quantity.value==''){
				alert('Please specify the Quantity');
				item.Quantity.focus;
				return false;			
		}
		
	if(item.Ref.value==''){
				alert('Please specify the Reference');
				item.Ref.focus;
				return false;			
		}
		item.submit;
	}
	
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
		
				
		//==========BEGIN REMOVE OPTION/S FROM DROP-DOWN FUNCTION ON THE FLY=====
		function RemoveOptions(Field){		   
		   if (Field.length==0) return;
		  
		   for (loop=Field.length - 1; loop >= 0; loop--) {
		       var GoneOption = Field.options[loop]		  
		       Field.remove(GoneOption.index);		        
		       }
		   
		 }

		//==============END REMOVE OPTION/S FUNCTION====================

		//var currentEntityType = <%=currentEntityType%>
		
		
		var totalContractAmt = 0;
	
		function UpdateVoucherAmount(inAmount, inAction){			

				
			switch(inAction){
				case 0:
					totalContractAmt = totalContractAmt - parseFloat(inAmount)			
					break;
				default:
					totalContractAmt = totalContractAmt + parseFloat(inAmount)				
					break;
						
			}
			
			document.all.item("txtTotal").value = FormatNum(totalContractAmt); 
			
		}
		
	function evaluateEntity(Val, Entity)
	{
      //Enable Printing if entity is Broker or Client
	  if (Val == 1 || Val == 3 || Val == 5){
	    document.getElementById("cmdPrint").style.display = ""
	  }
	  else {
	    document.getElementById("cmdPrint").style.display = "none"
	  }
	  
	  //Furher Processing
	  FetchAccounts(Entity)
	}
	
	function FormatNumber2(Obj){
	 var theVal = Obj.value
	 Obj.value = formatNum(theVal)
	}
	
	function FormatNumber(Obj,decimals){
	 var theVal = Obj.value
	 Obj.value = format_number2(theVal,decimals)
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
</head>
<body Class="Dialog">
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frm<%=DataSource%>", "txtDate","cmdDate","<%=FormatDate(rs.Fields("TradeDate"))%>",1);
</SCRIPT>

<form name = 'frm<%=DataSource%>' method = 'post' action = 'EditBondTrade.asp' id = "frm<%=DataSource%>" >
<table border="0" width="100%" cellspacing="1" cellpadding="1">
 <tr>
    <td width="15%">Code</td>
    <td width="54%"><input type="text" name="code" id="code" value="<%=rs.Fields("Client_DPA_")%>" size="10" onBlur="txtval = this.value; selectItem(client);"></td>
  </tr>
  <TR>
    <TD width="20%">Client</TD>
    <TD width="80%"><select size="1" name="client" id="client" onchange="UpdateClientCode();"
    onfocus="txtval = '';inputIsItemCode = 1;"
    onblur="txtval = '';inputIsItemCode = 1;" readonly>
    <option value="" SearchCode = '' SearchText = '' ></option>
    <%
		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [Clientlist] Order By ClientName"
        Set rst = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rst.EOF Or rst.BOF) Then
                rst.MoveFirst
                Do Until rst.EOF
				 if	Cdbl(rst.Fields("Client_DPA_")) = Cdbl(rs.Fields("Client_DPA_")) then	
				 %>
				 <option selected SearchCode =  '<%=rst.Fields("Client_DPA_")%>' SearchText = '<%=rst.Fields("ClientName")%>' value = '<%=rst.Fields("Client_DPA_")%>'><%=mid(rst.Fields("ClientName"),1,30)%></option>
				 <%  
                 else
                 %>
				 <option SearchCode = '<%=rst.Fields("Client_DPA_")%>' SearchText = '<%=rst.Fields("ClientName")%>' value = '<%=rst.Fields("Client_DPA_")%>'><%=mid(rst.Fields("ClientName"),1,30)%></option>
				 <%
                 end if 
                rst.MoveNext
                Loop
        End If
        rst.close
%>
    </select></TD></TR>
  <tr>
    <td width="15%">Bond Type</td>
    <td width="54%"><select name = 'cboBondType' id = 'cboBondType' size="1" onChange='FetchIssues(this);'>
     <%
        Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [Security] where OrderSecType_DPA_ = 1 order by Security_DPA_"
        Set rst = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	   
    %><option value="0"></option>
		<% while not rst.eof 
		   if rs.Fields("Security_DPA_") = rst.Fields("Security_DPA_") then
				   %>
				   <option value = '<%=rst.Fields("Security_DPA_")%>' selected><%=rst("SecurityCode")%></option><%  
		   else%>
				   <option value = '<%=rst.Fields("Security_DPA_")%>'><%=rst("SecurityCode")%></option><%  
		   end if
				 
		rst.movenext
		wend
		rst.close
     %>
    	 </select></td>
  </tr>
 
  <tr>
    <td width="15%">Bond Issue</td>
    <td nowrap>
    <select Name="cboBondIssue" id="cboBondIssue">	
    <%
        Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT  * From [IssueList] WHERE Security_DPA_ =" & rs.Fields("Security_DPA_")  & " Order By SecurityCode"        
        Set rst = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
	   
    %><option value=""></option>
		<% while not rst.eof 
		   if rs.Fields("Bond_DPA_") = rst.Fields("Bond_DPA_") then
				   %>
				   <option value = '<%=rst.Fields("Bond_DPA_")%>' selected><%=rst("BondIssue")%></option><%  
		   else%>
				   <option value = '<%=rst.Fields("Bond_DPA_")%>'><%=rst("BondIssue")%></option><%  
		   end if
				 
		rst.movenext
		wend
		rst.close
     %>	
    </select>
    </td>
  </tr>
		  <tr id="orderRow">
		    <td width="15%">Trade Type</td>
		    <td width="54%"><select name = 'TradeType' id = "TradeType" size="1">
		    <option selected value = ''></option>
		    <%
		    Set conn = GetActiveConnection("KBroker")
			    sqlStr = "SELECT * FROM [OrderType] order by OrderTypeDescription ASC"
			    Set rst = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
			   
			%>
				<% while not rst.eof 
				
			if rs.Fields("OrderType_DPA_") = rst.Fields("OrderType_DPA_") then
				   %>
				   <option value = '<%=rst.Fields("OrderType_DPA_")%>' selected><%=rst("OrderTypeDescription")%></option><%  
		   else%>
				   <option value = '<%=rst.Fields("OrderType_DPA_")%>'><%=rst("OrderTypeDescription")%></option><%  
		   end if
				rst.movenext
				wend
				rst.close
		    %>
		    </select>
		    </td>
		  </tr>
		  <tr >
		    <td width="15%" valign="top">Trade Date</td>
		    <td ><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
		  </tr>
		  
		  <tr >
		    <td width="15%" valign="top">Price</td>
		    <td width="5"><input type ="text" Name ="Price"  id="Price" size="20" value="<%=formatnumber(rs.Fields("Price"),2)%>" onChange="Javascript: FormatNumber(this,4)"></td>
		  </tr>
  <tr>
    <td width="15%">Quantity</td>
    <td width="54%"><input type ="text" Name ="Quantity"  id="Quantity" size="20" value="<%=formatnumber(rs.Fields("Quantity"),2)%>" onChange="Javascript: FormatNumber2(this)"></td>
  </tr>
  <tr>
    <td width="15%">Reference</td>
    <td width="54%"><textarea name="Ref" id ="Ref"><%=rs.Fields("Reference")%></textarea></td>
  </tr>
  <tr>
	  <td width="100%" colspan=2 align="center" valign=absBottom>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdSave' id = 'cmdSave' value=" Save ">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
	</td>
  </tr>
</table>

</form>
</body>

</html>