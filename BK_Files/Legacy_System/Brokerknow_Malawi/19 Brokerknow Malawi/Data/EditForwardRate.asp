<!--#include virtual="libroutines.asp"-->
<%
	const UDLName = "KBroker"
	const DataSource = "ForwardRatesList"
	const DataEntity = "ForwardRate"
	const DataEntityPlural = "ForwardRates"
	const ActionFolder = "Data"
	
	Dim UserId
	Dim action
	Dim conn 
	Dim sqlStr
	Dim rs
	Dim guid
	Dim guidStr
	Dim ID
	
	UserId=Session("UserID")
	action = ucase(Request.Form("action"))
	
	ID = Request("ID")
  
		If Trim(ID) = "" Then%>
                <script language = 'vbscript'>
                		ShowMessage "No record specified for editing"
                		
                </script>
                <% WriteDialogRefuseOpenScript
                response.end
        End If
	
	if action = "EXECUTE" then
	        
		    Dim ForwardRate
		
			ForwardRate = Request.Form("ForwardRate")
			ActivationDate = Request.Form("txtDate")
			ForwardRateID = Request.Form("ForwardRateID")
		
		    if  ActivationDate = "" then
			 %><script language = 'Javascript'>
					alert('Please specify the Activation Date.')
					window.history.back()
				</script><% response.end
			end if 
			
			if  not isdate(ActivationDate) then
			 %><script language = 'Javascript'>
					alert('Invalid Activation Date.')
					window.history.back()
				</script><% response.end
			end if 
			
			if  ForwardRate = "" then
			 %><script language = 'Javascript'>
					alert('Please specify the Forward Rate.')
					window.history.back()
				</script><% response.end
			end if 
			
			if not  isnumeric(ForwardRate) then
			 %><script language = 'Javascript'>
					alert('Invalid Forward Rate.')
					window.history.back()
				</script><% response.end
			end if 
			
			'save data
				Set conn = GetActiveConnection("KBroker")
												      
				 sqlStr = "Update [ForwardRate] Set ForwardRate = " & ForwardRate & ", ActivationDate = #" & ActivationDate & "#, " & _
									"ModifiedBy = " & session("UserID") & " , DateModified = #" & FormatDate(now()) & "#" & _
									" Where ForwardRate_DPA_ = " & ForwardRateID
									
				sqlStr = SQLServerFormatWithCustomMax(HandleQuote(sqlStr))
          
				conn.BeginTrans
					conn.Execute sqlStr
				conn.CommitTrans
					
				WritefraEnabledDialogCloseScript
		
				conn.Close
			    Set conn = Nothing
					
				Response.End	
   	End if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add <%=DataEntity%></title>

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
<%
Set conn = GetActiveConnection("KBroker")
       
	sqlStr = "SELECT * from [ForwardRatesList] where Bond_DPA_  = " & ID

        Set rs = conn.Execute(SQLServerFormat(HandleQuote(sqlStr)))
        If rs.EOF Or rs.BOF Then%>
                <script language = 'vbscript'>
                		ShowMessage "The selected Bond cannot be retrieved!"
                		
                </script>
                <% WriteDialogRefuseOpenScript
                response.end
        End If
         
 if isnull(rs.fields("ForwardRate_DPA_")) then
 %>
                <script language = 'vbscript'>
                		ShowMessage "There is no Forward Rate to be Edited."
                		
                </script>
                <% WriteDialogRefuseOpenScript
                response.end
 end if


%>
<div id="spiffycalendar" class="text" STYLE="z-Index: 1000"></div>

<SCRIPT language="JavaScript">
	var cal=new ctlSpiffyCalendarBox("cal", "frm<%=DataSource%>", "txtDate","cmdDate","<%=FormatDate(rs.Fields("ActivationDate"))%>",1);
</SCRIPT>

<form name = 'frm<%=DataSource%>' method = 'post' action = 'EditForwardRate.asp' id = "frm<%=DataSource%>" >

<table border="0" width="100%" cellspacing="1" cellpadding="1">
 <tr>
    <td width="15%">Bond Type</td>
    <td width="54%"><input type="text" name="BondType" id="BondType" value="<%=rs.fields("BondType")%>" size="20" class="readonly"></td>
  </tr>
 <tr>
    <td width="15%">Bond Issue</td>
    <td width="54%"><input type="text" name="BondIssue" id="BondIssue" value="<%=rs.fields("BondIssue")%>" size="20" class="readonly"></td>
  </tr>
  <tr>
    <td width="15%">Issue Date</td>
    <td width="54%"><input type="text" name="IssueDate" id="IssueDate" value="<%=formatdate(rs.fields("IssueDate"))%>" size="20" class="readonly"></td>
  </tr>
  <tr>
    <td width="15%">Maturity Date</td>
    <td width="54%"><input type="text" name="MaturityDate" id="MaturityDate" value="<%=formatdate(rs.fields("MaturityDate"))%>" size="20" class="readonly"></td>
  </tr>
  <tr>
    <td width="15%">Face Value</td>
    <td width="54%"><input type="text" name="FaceValue" id="FaceValue" value="<%=formatnumber(rs.fields("FaceValue"),2)%>" size="20" class="readonly"></td>
  </tr>
  <tr>
    <td width="15%">Forward Rate</td>
    <td width="54%"><input type="text" name="ForwardRate" id="ForwardRate" value="<%=rs.fields("ForwardRate")%>" size="20"></td>
  </tr>
  <tr>
    <td width="15%">Activation Date</td>
    <td width="54%"><SCRIPT language="JavaScript">cal.writeControl();</SCRIPT></td>
  </tr>
  <tr>
	  <td width="100%" colspan=2 align="center" valign=absBottom>
		<BR>
		<input type = 'submit' Class=Buttons name ='cmdSave' id = 'cmdSave' value=" Save ">
    	<input type = 'button' Class=Buttons name ='cmdCancel' id = "cmdCancel" value="Cancel" onclick = "JavaScript: window.self.close()">
		<input type = 'hidden' name ='action' id = 'action' value="Execute">
		<input type = 'hidden' name ='ID' id = 'ID' value="<%=ID%>">
		<input type = 'hidden' name ='ForwardRateID' id = 'ForwardRateID' value="<%=rs.fields("ForwardRate_DPA_")%>">
	</td>
  </tr>
</table>

</form>
</body>

</html>