<html>

<head>
	<meta http-equiv="Content-Language" content="en-uk">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	
	<title>Trade Affirmation</title>
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/default.css"> 
	<LINK REL="STYLESHEET" TYPE="TEXT/CSS" HREF="../STYLE/webparts.css">
	 <SCRIPT language=Javascript src="../scripts/common.js"></SCRIPT> 
	 <SCRIPT language=Javascript src="../scripts/fhsupport.js"></SCRIPT>
	  <!--CALENDAR -->
	<link rel="stylesheet" type="text/css" href="CALENDAR/calendar.css">
	<SCRIPT language=Javascript src="Calendar/calendar.js"></SCRIPT>
	<SCRIPT language=VBScript src="../scripts/reports.vbs"></SCRIPT>
	<SCRIPT language=Javascript src="../scripts/reports.js"></SCRIPT>
	<style media="print">
		@page {
			@top{font-family: Helvetica, Arial, sans-serif;
				font-size: 150%;
				font-weight: bolder;
				text-align: left;
				content: "<%= FormatDate(Date) %>";			
			}		
			
			margin-left: 2cm;
			margin-right: 5cm;
			margin-top: 1cm;    
			margin-bottom: 2cm;
			size: portrait;
			
			br.newpage{
				page-break-before:always;
			}
			
			tr.pageNumbering{
				display:none;
			}
			
		}

	</style>
</head>

<body Class="Reports">

<!--#include file="../libroutines.asp"-->


<%

server. scripttimeout = 1000000
genReport = Request.Form("genReport")
selectedTradeDate = Request.Form("txtDate")
useOwner = Request.Form("useOwner")
selOwners = Request.Form("selOwners")
useClient = Request.Form("useClient")
selClient = Request.Form("cboClient")

If genReport <> "1" Or Not IsDate(selectedTradeDate) Then%>
	<Script Language="JavaScript">
		report_SetBodyClass();
	
		function validateForm(frm){			
			if (frm.txtDate.value==''){
				alert("Select a date");
				frm.txtDate.focus();
				return;
			}
			
			frm.target = '_self';			
			frm.submit();
		}
		function switchDisplay(obj){
			if (obj.style.display=='none') obj.style.display = '';
			else obj.style.display = 'none';
		}
		
		
	function calendarChange(dateValue){
			var Input = document.all.item('AllSelOwners');
			var Output =  document.all.item('selOwners');
			var dateVal;
			var mainObj = document.all.item('txtDate');
			
			if (dateValue==null || dateValue=='undefined') dateVal = mainObj.value;
			else dateVal = dateValue;
			
			Output.length = 0;
		     for (loop=0; loop < Input.length; loop++){
		     		if (Input.options[loop].TAG == clientFormatDate(dateVal)){		     			
		     		    NewOption = new Option();   			    
		   			    NewOption.text = Input.options[loop].text;
		   			    NewOption.value = Input.options[loop].value;			
		   			    Output.add(NewOption, 0);
		     		}	
		     }
		}
		
		function LoadClient(accountno, element)
		{
		 var clientcode = document.frmMain.elements("txtClientCode").value
		 var clientcds = document.frmMain.elements("txtcdsno").value
		 var clientname = document.frmMain.elements("txtclientname").value
		 var clientcobo = document.getElementById("cboClient");		 
		 
	     var x_clientname;
			
		 if (element == 'txtClientCode')
		 {
			clientcds = ''
			clientname = ''
			
			if (clientcode == '')
			{
			document.frmMain.elements("txtClientCode").value = 'Code'
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			return;
			}
			
		 }
		 else if (element == 'txtcdsno')
		 {
			clientcode = ''
			clientname = ''

			if (clientcds == '')
			{
			document.frmMain.elements("txtcdsno").value = 'CDS No.'
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			return;
			}
						
		 }
		 else if (element == 'txtclientname')
		 {
			clientcode = ''
			clientcds = ''

			if (clientname == '')
			{
			document.frmMain.elements("txtclientname").value = 'Client Name';
			clientcobo.length = 1;
			clientcobo[0].text = 'Load Account';
			clientcobo[0].value = '';
			return;
			}
			
		 }

				xmlhttp = createXMLHTTPObj();
				
				url="GetList.asp?clientcode="+clientcode+"&cdsno="+clientcds+"&clientname="+clientname+"&action=SLoadClient";
				
				//alert(url);

				xmlhttp.open("GET",url,true);

				xmlhttp.onreadystatechange=function() 
				  {
							if (xmlhttp.readyState==4) 
							{
								returnStr = xmlhttp.responseText;
								returnStr = getBodyHTML(returnStr);
								
								//alert(returnStr);

								myArray = returnStr.split(";");
								
								x_clientname = myArray[1]

								if (x_clientname.length > 12) 
								{
									x_clientname = x_clientname.substring(0,16)  + '...';
								}
								
								//document.getElementById("cboAccount").options.length = 0;
								clientcobo.length = 1;
								if (element != 'txtclientname')
								{
									document.frmMain.elements("txtClientCode").value = myArray[0];
									document.frmMain.elements("txtcdsno").value =myArray[2];
									document.frmMain.elements("txtclientname").value = x_clientname;
									
									clientcobo[0].SearchCode = myArray[0];
									clientcobo[0].SearchText = myArray[1];
									
									clientcobo[0].text = myArray[1];
									clientcobo[0].value = myArray[0];
									
								}
								else
								{

									var myArrayx;
									var myArrayz;
									
									//alert(returnStr);
									myArrayx = returnStr.split("|");
									myArrayxsize = myArrayx.length - 1;
									
									//alert(myArrayxsize);

									for (i=myArrayxsize; i>=0; i--)
									{
										
										myArrayz = myArrayx[i].split(";");

										//alert(myArrayz)
										
										document.frmMain.elements("txtClientCode").value = '';
										document.frmMain.elements("txtcdsno").value = '';
										document.frmMain.elements("txtclientname").value = '';
										
										document.getElementById("cboClient").options[i] = new Option(myArrayz[2],myArrayz[0]);
																				
									}
									
								}
							}
					}
				 
				 xmlhttp.setRequestHeader('Accept','message/x-jl-formresult');
				 xmlhttp.send();
		}
		
		function ClearFields(element)
		{
		   if (element == 'txtClientCode')
		   {
			document.frmMain.elements("txtClientCode").value = ''
			document.frmMain.elements("txtcdsno").value = 'CDS No.'
			document.frmMain.elements("txtclientname").value = 'Client Name'
			return;
		   }
		   if (element == 'txtcdsno')
		   {
			document.frmMain.elements("txtClientCode").value = 'Code'
			document.frmMain.elements("txtcdsno").value = ''
			document.frmMain.elements("txtclientname").value = 'Client Name'
			return;
		   }
		   if (element == 'txtclientname')
		   {
		    document.frmMain.elements("txtclientname").value = ''
			document.frmMain.elements("txtClientCode").value = 'Code'
			document.frmMain.elements("txtcdsno").value = 'CDS No.'
			return;
		   }		
		   
		}

		function updatefields(selectedclient)
		{
		 document.frmMain.elements("txtclientname").value = '';
		 document.frmMain.elements("txtcdsno").value = '';
		 document.frmMain.elements("txtClientCode").value = selectedclient.value;
		}
		
		var cal = new ctlSpiffyCalendarBox("cal", "frmMain", "txtDate","cmdDate","<%= FormatDate(Date) %>",1);
		
		
	</Script>
	
	<Script Language="VBScript">
		Function clientFormatDate(theDate)
			On Error Resume Next
			clientFormatDate = Day(theDate) & "-" & MonthName(Month(theDate), True) & "-" & Year(theDate)
			If Err.Number > 0 Then
				clientFormatDate = theDate
			End If
		End Function
	</Script>
	
	<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>
	<form method="POST" action="TradeAffirmationQuick.asp" Name="frmMain" id="frmMain">
		<input type="hidden" value="1" name="genReport">	
		<input type="hidden" Name="timeLimit" value="1">
		
  <table>
    <tr> 
      <td colspan="2"> <input type="checkbox" OnClick="JavaScript: switchDisplay (document.all.item('ownerSelectRow')); " class="BorderLess" name="useOwner" id="useOwner" value="1"> 
        &nbsp; &nbsp; <label for="useOwner" style="cursor: hand">Narrow down to 
        specific owner/s (optional)</label></td>
    </tr>
    <tr style="display: none;" id="ownerSelectRow" align="right"> 
      <td colspan=2> <select name="AllSelOwners" style="display: none">
          <%
						Set conn = GetActiveConnection("KBroker")
						sqlStr = "SELECT DISTINCT LotTDate, Owner FROM traff_TradeAffirmation"   
						Set Rs = Conn.Execute (sqlStr)
						If Not (Rs.EOF Or Rs.BOF) Then
							Do Until Rs.EOF%>
          <option TAG="<%= FormatDate(Rs.Fields("LotTDate").value) %>" value="<%= Rs.Fields("Owner").Value %>"><%= Rs.Fields("Owner").Value %></option>
          <%Rs.MoveNext
							Loop
						End If%>
        </select> 
        <!--<select name="selOwners" size="10" multiple>
											
					</select>-->
        <Script Language="JavaScript">
						calendarChange ("<%= FormatDate(Date) %>")						
					</Script> </td>
    </tr>
    <tr> 
      <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Select 
        any day of month &nbsp;&nbsp;&nbsp;&nbsp; 
        <script language="JavaScript">
						cal.writeControl();
					</script></td>
      <td width="4%" align="left">&nbsp; </td>
    </tr>
    <tr> 
      <td> <input type="checkbox" OnClick="JavaScript: switchDisplay (document.all.item('clientSelectRow')); " class="BorderLess" name="useClient" id="useClient" value="1"> 
        &nbsp; &nbsp; <label for="useClient" style="cursor: hand">Narrow down 
        to specific client (optional)</label></td>
      <td>&nbsp;</td>
    </tr>
    <tr style="display: none;" id="clientSelectRow" align="right"> 
      <td nowrap align="left" valign="bottom" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type = 'text' name ='txtClientCode' id = 'txtClientCode3' size="5" value = 'Code' onBlur="LoadClient(cboClient, this.name);"  onClick  = "ClearFields(this.name)"> 
        &nbsp; <input type = 'text' name ='txtcdsno' id = 'txtcdsno' size="10" value = 'CDS NO.' onBlur="LoadClient(cboClient, this.name);" onClick  = "ClearFields(this.name)"> 
        <input type = 'text' name ='txtclientname' id = 'txtclientname' size="15" value = 'Client Name' onBlur="LoadClient(cboClient, this.name);" onClick  = "ClearFields(this.name)"> 
        <select name = 'cboClient' id = "cboClient" size="1" onChange ="updatefields(this);">
          <option SearchCode = "" SearchText = "" value = '' >Load Account</option>
        </select> </td>
    </tr>
    <tr> 
      <td colspan=2><input type="Button" class="Buttons" OnClick="JavaScript: validateForm(document.all.item('frmMain'))" Value=" Generate... ">
        &nbsp;&nbsp; </td>
    </tr>
  </table>
		
	</form>
		
	<%Set rs = Nothing
	Set Conn = Nothing
	Response.End
End If

%>

<% DrawPageFunctions True, True, True %>

<%


   Dim conn 
   Dim sqlStr
   Dim rs
   Dim isDifferentBroker
   Dim currBrokerCode
   
   Set conn = GetActiveConnection("KBroker")
   conn.commandTimeout = 0


   conn.execute("traff_PopulateTradeAffirmationData '" & FormatDate(selectedTradeDate) & "'")
   

   sqlStr = "SELECT Order_DPA_ FROM traff_TradeAffirmation WHERE LotTDate = '" & FormatDate(selectedTradeDate) & "'"


   If useOwner = "1" Then
		
		If selOwners <> "" Then
			selOwners = Replace(selOwners, "'", "''")
			selOwners = Replace(selOwners, ",", "','")
			selOwners = "'" & selOwners & "'"		
			sqlStr = sqlStr & " AND Owner IN (" & selOwners & ")"
		End If	
   End If
   
   If useClient = "1" Then
		
		If selClient <> "" Then		
			sqlStr = sqlStr & " AND Client_DPA_ = " & selClient
		End If	
   End If
   'Response.Write sqlStr : Response.end
   Set Rs = Conn.Execute (sqlStr)
   If Rs.EOF Or Rs.BOF Then%>
		<Script Language="JavaScript">	
			ShowMessage('There were no traded items under the selected date');
			window.parent.history.go(-1)
		</Script>
		<%Set Conn = Nothing
		Set Rs = Nothing
		Response.End
   End If
   
   
   
   
   'sqlStr = "SELECT Order_DPA_, SUM(LotQty) As Quantity, SUM(SettlementAmount) As SettlementAmount, OrderRef FROM traff_TradeAffirmation WHERE LotTDate = '" & FormatDate(selectedTradeDate) & "'"

   sqlStr = "SELECT   Order_DPA_, OrderRef FROM traff_TradeAffirmation WHERE LotTDate = '" & FormatDate(selectedTradeDate) & "'"
    
   If useOwner = "1" Then
		If selOwners <> "" Then
			sqlStr = sqlStr & " AND Owner IN (" & selOwners & ")"
		End If	
   End If
   
    If useClient = "1" Then
		If selClient <> "" Then		
			sqlStr = sqlStr & " AND Client_DPA_ = " & selClient
		End If	
   End If
   sqlStr = sqlStr & "	GROUP BY OrderRef, Order_DPA_ "

   'response.write sqlStr
   'response.end
   
   Set sumRs = Conn.Execute(sqlStr)         
   Dim pageNumber
	
	pageNumber = 0
	
   Do Until sumRs.EOF
   response.flush
		'totalQty = sumRs.Fields("Quantity").Value ' Calculate total rather than Quering
		'settlementAmt = sumRs.Fields("SettlementAmount").Value
		OrderRef = sumRs.Fields("OrderRef").Value	
				 
		sqlStr = "SELECT  * FROM traff_TradeAffirmation WHERE LotTDate = '" & FormatDate(selectedTradeDate) & "' And OrderRef = '" & OrderRef & "' AND Order_DPA_ = " & sumRs.Fields("Order_DPA_").Value		
		
		'Response.Write(sqlstr)
		'Response.End()
		
		Set Rs = Conn.Execute (sqlStr)
		OwnerName = Rs.Fields("Owner").Value
		if isnull(orderRef) or trim(orderRef) = "" then
				orderRef = sumRs.Fields("Order_DPA_").Value
		else
				orderRef = sumRs.Fields("Order_DPA_").Value & "/" & orderRef
		end if
		pageNumber = pageNumber + 1
%> 

<table border="0" cellspacing=2 cellpadding=2 class="ReportsTable" width="100%">
	<tr class="pageNumbering">
		<td align="left" >
			<FONT FACE=ARIAL SIZE=2><B>Page <%=pageNumber%>	</B></FONT>	
		</td>		
	</tr>
	<tr>
		<td align="center" height="250px">
			<Img Src="../data/photos/aaprintlogo.jpg">			
		</td>		
	</tr>
</table>	  

  <table border="0" cellspacing="0" cellpadding="5" style="font-family: Arial" width="100%">

    <tr>
      <td colspan="2" align="center"><b><font size="4"><u>TRADE
        AFFIRMATION</u></font></b></td>
    </tr>
    <tr>
      <td width="20%">To:</td>
      <td><%= OwnerName %></td>
    </tr>
    <tr>
      <td width="20%">Your Account:</td>
      <td><%= Rs.Fields("Account").Value %></td>
    </tr>
    <tr>
      <td width="20%"></td>
      <td><%= Rs.Fields("AccountAddress").Value %></td>
    </tr>
    <tr>
      <td width="20%">Fax Number:</td>
      <td><%= Rs.Fields("ClientFax").Value %></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td width="20%"><b><u>Re:</u><b></td>
      <td><b><u>
		<%= Rs.Fields("ReferenceHeader")  %>
      </u></b></td>
    </tr>
   
    <tr>
      <td>Order Ref:</td>
      <td><%= OrderRef %></td>
    </tr>
    <tr>
      <td>Trade Date:</td>
      <td><%= FormatDate(selectedTradeDate) %></td>
    </tr>
	<% 'Calculate total quantity
          totalQty = 0
		  settlementAmt = 0
		  
          Do Until Rs.EOF
		  
		   totalQty = totalQty + Rs.Fields("LotQty").Value
		   settlementAmt = settlementAmt + Rs.Fields("SettlementAmount").Value
		    Rs.MoveNext
          Loop
		  Rs.movefirst
          %>
    <tr>
      <td>Total Quantity</td>
      <td><b><%= FormatNumCommasOnly(totalQty) %></b></td>
    </tr>
    <tr>
      <td>Settlement Amount:</td>
      <td><b><%= FormatNum(settlementAmt) %></b></td>
    </tr>
    <tr>
      <td>Settlement Date:</td>
      <td><% '= FormatDate(upperLimitDate) %></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="2">Breakdown of slip details:</td>
    </tr>
    <tr>
    <td>&nbsp;</td>
      <td>
      	
        <table border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td>Slip No</td>
            <td>Contract</td>
            <td align="right">Quantity</td>
            <td align="right">Deal Price</td>
            <td align="right">Settle</td>
            
          </tr>
          <%
          totalQty = 0

          Do Until Rs.EOF
		  'response.flush
		  if cint(Rs.Fields("OrderSecType_DPA_").Value) = 1 then
			  priceFormatted = FormatNumber(Rs.Fields("LotPrice").Value, 4)
		  else
			  priceFormatted = FormatNumber(Rs.Fields("LotPrice").Value, 2)
		  end if
			
		  %>
          
          <tr>
            <td><%= Rs.Fields("LotSlipNo").Value %></td>
            <td><%= Rs.Fields("ContractNumber").Value %></td>
            <td align="right"><%= FormatNumCommasOnly(Rs.Fields("LotQty").Value) %></td>
            <td align="right"><%= priceFormatted  %></td>
            <td align="right"><b><%= FormatNum(Rs.Fields("SettlementAmount").Value) %></b></td>
            
          </tr>
          
          <%
		   totalQty = totalQty + Rs.Fields("LotQty").Value
		  
		  Rs.MoveNext
          Loop
          %>
        </table>
      </td>
    </tr>
  </table>
<br>  
<table border="0" cellspacing=2 cellpadding=2>
	<tr>
		<td align="left">
<PRE><font face="Arial"><b>		
for African Alliance Malawi Securities</b>
</font></PRE>			
		</td>		
	</tr>
</table>  

<%		sumRs.MoveNext
		'important!		
			If Not sumRs.EOF Then %>
				<BR class="newpage">
		<%	End If
	Loop

	Set sumRs = Nothing
	Set Rs = Nothing
Set Conn = Nothing%>
</body>

</html>
