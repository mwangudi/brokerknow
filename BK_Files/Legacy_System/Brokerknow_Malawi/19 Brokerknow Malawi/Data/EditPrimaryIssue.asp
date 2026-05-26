<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "PrimaryIssuesList"
	const DataEntity = "PrimaryIssue"
	const DataEntityPlural = "PrimaryIssues"
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
				
		    Dim code
		    Dim client
			Dim BondType
			Dim BondIssue
			Dim ValueDate
		    Dim FaceValue
			Dim Price
			Dim Rate
			Dim AppNo
			Dim AccAmount
		    Dim PaymentType
	        Dim PaymentDate
	        Dim Reference
	        Dim PaidAmount
	        Dim Narrative
	        
			code = Request.Form("code")
			client = Request.Form("client")
			BondType = Request.Form("cboBondType")
			BondIssue = Request.Form("cboBondIssue")
			ValueDate = Request.Form("txtValueDate")
			FaceValue = Request.Form("FaceValue")
			Price = Request.Form("Price")
			AppNo = Request.Form("ApplicationNo")
			AccAmount = Request.Form("AcceptedAmount")
			PaymentType = Request.Form("PaymentType")
			PaymentDate = Request.Form("txtPDate")
			Reference = Request.Form("Reference")
			PaidAmount = Request.Form("PAmount")
			Narrative = Request.Form("Narrative")
			
			'Validations  
			if  trim(AppNo) = "" then
			 %><script language = 'vbscript'>
                ShowMessage "Please specify the Application No"
                window.self.close
                </script>
                <% response.end
			end if  
            
			If trim(code) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Client Code"
                window.self.close
                </script>
                <% response.end
            End If
			
			If trim(client) = "" or trim(client) = 0 Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Client."
                window.self.close
                </script>
                <% response.end
            End If
           
			If trim(BondType) = "" or trim(BondType) = 0 Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Bond Type."
                window.self.close
                </script>
                <% response.end
            End If
			
			If trim(BondIssue) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Bond Issue."
                window.self.close
                </script>
                <% response.end
            End If 
			
			If trim(ValueDate) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Value Date."
                window.self.close
                </script>
                <% response.end
            End If
            
			If trim(FaceValue) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Face Value."
                window.self.close
                </script>
                <% response.end
            End If
           
			If not isnumeric(trim(FaceValue)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Invalid Face Value."
                window.self.close
                </script>
                <% response.end
            End If
		
            If trim(Price) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Price."
                window.self.close
                </script>
                <% response.end
            End If
           
			If not isnumeric(trim(Price)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Invalid Price."
                window.self.close
                </script>
                <% response.end
            End If
            
			If trim(AccAmount) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Accepted Amount."
                window.self.close
                </script>
                <% response.end
            End If
           
			If not isnumeric(trim(AccAmount)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Invalid Accepted Amount."
                window.self.close
                </script>
                <% response.end
            End If
           
			If trim(PaymentType) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Payment Type."
                window.self.close
                </script>
                <% response.end
            End If
			
			If trim(PaidAmount) = "" Then%>
                <script language = 'vbscript'>
                ShowMessage "Please specify the Paid Amount."
                window.self.close
                </script>
                <% response.end
            End If
           
			If not isnumeric(trim(PaidAmount)) Then%>
                <script language = 'vbscript'>
                ShowMessage "Invalid Paid Amount."
                window.self.close
                </script>
                <% response.end
            End If
            
			'save data
					
		     Set conn = GetActiveConnection("KBroker")
				 conn.BeginTrans		
								      
				    sqlStr = "Update [PrimaryIssues] set Client_DPA_ = " & client & ", Bond_DPA_ =" & BondIssue & ", " & _
									"ValueDate = #" & FormatDate(ValueDate) & "#, FaceValue = " & Ccur(FaceValue) & ", " & _
									"Price = " & Ccur(Price) & ", ApplicationNo = '" & AppNo & "', " &_
									"AcceptedAmount = " & Ccur(AccAmount) & ", PaymentTypes_DPA_ = " & PaymentType & " , PaymentDate = #" & FormatDate(PaymentDate) & "#, " &_
									"Reference = '" & Reference & "', PaidAmount = " & Ccur(PaidAmount) & " , " &_
									"Narrative = '" & Narrative & "', ModifiedBy = " & session("UserID") & " ,DateModified = #" & FormatDate(now()) & "# Where PrimaryIssues_DPA_=" & ID
									
					sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
                      
					conn.Execute sqlStr
					conn.CommitTrans
					
					WritefraEnabledDialogCloseScript
		
				    conn.Close
					Set conn = Nothing
					
					Response.End	
   	End if
   	
 Set conn = GetActiveConnection("KBroker")
	sqlStr = "SELECT * FROM PrimaryIssuesList WHERE PrimaryIssues_DPA_=" & ID
			        
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
	
	function FormatNumber2(Obj){
	 var theVal = Obj.value;
	 Obj.value = formatNum(theVal);
	}
	
	function UpdateAmount(){
	 var frm = document.frm<%=DataSource%>  ;
	 var Price  = frm.Price.value ; 
	 var AcceptedAmount = frm.AcceptedAmount.value ;
	
	 frm.PAmount.value = formatNum(Price*AcceptedAmount)
	}
	
	function FormatNumber(Obj,decimals){
	 var theVal = Obj.value; 
	 Obj.value = format_number2(theVal,decimals);
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
	var cal=new ctlSpiffyCalendarBox("cal", "frm<%=DataSource%>", "txtValueDate","cmdValueDate","<%=FormatDate(rs.Fields("valueDate"))%>",1);
	var cal1=new ctlSpiffyCalendarBox("cal1", "frm<%=DataSource%>", "txtPDate","cmdPDate","<%=FormatDate(rs.Fields("PaymentDate"))%>",1);
</SCRIPT>

<form name = 'frm<%=DataSource%>' method = 'post' action = 'EditPrimaryIssue.asp' id = "frm<%=DataSource%>" >
<table border="0" width="100%" cellspacing="1" cellpadding="1">
 <TR>
    <TD colspan="2" align="left"><b>APPLICATION DETAILS</b></TD>
   </TD></TR>
  <TR>
    <TD width="20%"></TD>
    <TD width="80%"></TD></TR>
 <TR>
    <TD width="20%">Application No</TD>
    <TD width="80%"><input type ="text" Name ="ApplicationNo"  id="ApplicationNo" size="20" Value="<%=rs.Fields("ApplicationNo")%>"></TD></TR>
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
  <TR>
    <TD width="20%">Face Value</TD>
    <TD width="80%"><input type ="text" Name ="FaceValue"  id="FaceValue" size="20" Value="<%=formatnumber(rs.Fields("FaceValue"),2)%>" onChange="Javascript: FormatNumber2(this);"></TD></TR>
  <TR>
    <TD width="20%">Price</TD>
    <TD width="80%"><input type ="text" Name ="Price"  id="Price" size="20" Value="<%=formatnumber(rs.Fields("Price"),4)%>" onChange="Javascript: UpdateAmount(); FormatNumber(this,4);"></TD></TR>
    <TD width="20%">Value Date</TD>
    <TD width="80%"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></TD></TR>
  <TR>
    <TD width="20%">Accepted Value</TD>
    <TD width="80%"><input type ="text" Name ="AcceptedAmount"  id="AcceptedAmount" size="20" Value="<%=formatnumber(rs.Fields("AcceptedAmount"),2)%>" onChange="Javascript: UpdateAmount(); FormatNumber2(this)"></TD></TR>
  <TR>
    <TD ></TD>
    <TD></TD></TR>
    <TR>
    <TD colspan="2" align="LEFT"><b>PAYMENT DETAILS<b></TD>
    </TR>
  <TR>
  <TR>
    <TD width="20%">Payment Type</TD>
    <TD width="80%"><select size="1" name="PaymentType" id="PaymentType">
    <option value="" ></option>
    <%
		Set conn = GetActiveConnection("KBroker")
        sqlStr = "SELECT * FROM [PaymentTypes] Order By Description"
        Set rst = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If Not (rst.EOF Or rst.BOF) Then
                rst.MoveFirst
                Do Until rst.EOF
		         if rst.Fields("PaymentTypes_DPA_") = rs.Fields("PaymentTypes_DPA_") then ' Default = Cheque				
				 %><option value = '<%=rst.Fields("PaymentTypes_DPA_")%>' selected><%=rst.Fields("Description")%></option><%  
				 else%>
				 <option value = '<%=rst.Fields("PaymentTypes_DPA_")%>'><%=rst.Fields("Description")%></option>
				 <%
				 end if
                rst.MoveNext
                Loop
        End If
        rst.close
%>
    </select></TD></TR>
  <TR>
    <TD></TD>
    <TD></TD></TR>
  
    <TD width="20%">Date od Payment</TD>
    <TD width="80%"><SCRIPT language="JavaScript">cal1.writeControl();</SCRIPT></TD></TR>
  <TR>
    <TD width="20%">Reference</TD>
    <TD width="80%"><textarea name="Reference" id ="Reference"><%=rs.Fields("Reference")%></textarea></TD></TR>
  <TR>
    <TD width="20%">Amount</TD>
    <TD width="80%"><input type ="text" Name ="PAmount"  id="PAmount" class="readonly" id="PAmount" size="20" readonly value="<%=formatnumber(rs.Fields("PaidAmount"),2)%>"></TD></TR>
  <TR>
    <TD width="20%">Narrative</TD>
    <TD width="80%"><textarea name="Narrative" id ="Narrative"><%=rs.Fields("Narrative")%></textarea></TD></TR>
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